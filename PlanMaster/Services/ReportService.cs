using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PlanMaster.Models;

namespace PlanMaster.Services;

public sealed class ReportService
{
    public string TemplatePath { get; }

    // ИМЕНА ЗАКЛАДОК В ШАБЛОНЕ
    private const string BookmarkTeachingTables = "BM_TEACHING_TABLES";
    private const string BookmarkSummaryTable = "BM_SUMMARY_TABLE";

    // <<< ВОТ ЭТА ЗАКЛАДКА ДОЛЖНА БЫТЬ У ТЕБЯ В ШАБЛОНЕ в разделе "Дополнительная нагрузка"
    // Если имя другое — поменяй тут.
    private const string BookmarkExtraTables = "BM_EXTRA_TEACHING_TABLES";

    // Компактный текст таблиц
    private const string FontSizeBody = "16";   // 8pt
    private const string FontSizeHeader = "16"; // 8pt
    private const string FontSizeTitle = "20";  // 10pt

    public ReportService(string templatePath)
    {
        TemplatePath = templatePath;
    }

public string GenerateDocx(
    string outputDocxPath,
    IReadOnlyList<PlanTable> tables,
    IReadOnlyList<SummaryRow> summaryRows)
{
    Directory.CreateDirectory(Path.GetDirectoryName(outputDocxPath)!);
    File.Copy(TemplatePath, outputDocxPath, overwrite: true);

    using var doc = WordprocessingDocument.Open(outputDocxPath, true);
    var body = doc.MainDocumentPart!.Document.Body!;

    // ВАЖНО: фиксируем порядок (как у тебя везде по SheetName)
    var ordered = tables
        .OrderBy(t => t.SheetName)
        .ThenBy(t => t.SemesterTitle)
        .ToList();

    // 1-я закладка: первые 2 таблицы
    var firstBlock = ordered.Take(2).ToList();

    // 2-я закладка: следующие 2 таблицы
    var secondBlock = ordered.Skip(2).Take(2).ToList();

    InsertAtBookmark(body, "BM_TEACHING_TABLES", BuildTeachingTables(firstBlock, isExtraLoad: false));
    InsertAtBookmark(body, "BM_EXTRA_TEACHING_TABLES", BuildExtraLoadBlock(secondBlock));

    // Итоговая
    InsertAtBookmark(body, "BM_SUMMARY_TABLE", new[] { BuildSummaryTable(summaryRows) });

    doc.MainDocumentPart.Document.Save();
    return outputDocxPath;
}


    public string ConvertToPdfWithLibreOffice(string docxPath, string outputDir)
    {
        Directory.CreateDirectory(outputDir);

        // На Windows/macOS/Linux будет работать, если LibreOffice установлен и soffice доступен в PATH
        var soffice = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "soffice.exe" : "soffice";

        var psi = new ProcessStartInfo
        {
            FileName = soffice,
            ArgumentList =
            {
                "--headless",
                "--convert-to", "pdf",
                "--outdir", outputDir,
                docxPath
            },
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using var p = Process.Start(psi)!;
        p.WaitForExit();

        var pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");

        if (!File.Exists(pdfPath))
        {
            var err = p.StandardError.ReadToEnd();
            throw new Exception("Не удалось сконвертировать в PDF через LibreOffice. " + err);
        }

        return pdfPath;
    }

    /// <summary>
    /// Кроссплатформенное открытие файла системным приложением.
    /// </summary>
    public void OpenFile(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("Файл не найден", path);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/c start \"\" \"{path}\"",
                CreateNoWindow = true,
                UseShellExecute = false
            });
            return;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "open",
                ArgumentList = { path },
                UseShellExecute = false
            });
            return;
        }

        // Linux
        Process.Start(new ProcessStartInfo
        {
            FileName = "xdg-open",
            ArgumentList = { path },
            UseShellExecute = false
        });
    }

    /// <summary>
    /// Печать PDF.
    /// Linux/macOS: пробуем lp.
    /// Windows: безопасно открываем PDF (печать в просмотрщике).
    /// Возвращает true, если отправили напрямую, false если открыли viewer.
    /// </summary>
    public bool PrintPdf(string pdfPath)
    {
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("PDF не найден", pdfPath);

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            OpenFile(pdfPath);
            return false;
        }

        var psi = new ProcessStartInfo
        {
            FileName = "lp",
            ArgumentList = { pdfPath },
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        };

        using var p = Process.Start(psi)!;
        p.WaitForExit();

        if (p.ExitCode != 0)
        {
            var err = p.StandardError.ReadToEnd();
            throw new Exception("Не удалось отправить на печать через lp: " + err);
        }

        return true;
    }

    // -------------------- helpers --------------------

private static void InsertAtBookmark(Body body, string bookmarkName, IEnumerable<OpenXmlElement> elements)
{
    var bookmarkStart = body.Descendants<BookmarkStart>()
        .FirstOrDefault(b => b.Name == bookmarkName);

    if (bookmarkStart == null)
        throw new Exception($"Закладка '{bookmarkName}' не найдена в шаблоне.");

    // Ищем абзац, в котором стоит закладка (таблицу внутрь Run/Paragraph нельзя)
    var anchor = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault()
                 ?? bookmarkStart.Parent
                 ?? throw new Exception("Некорректная закладка в шаблоне (Parent == null).");

    OpenXmlElement insertAfter = anchor;

    foreach (var el in elements)
    {
        insertAfter.InsertAfterSelf(el);
        insertAfter = el;
    }
}


    /// <summary>
    /// Определяем "дополнительную нагрузку" по названию листа/таблицы.
    /// Подстрой под твои реальные имена листов.
    /// </summary>
    private static bool IsExtraLoadTable(PlanTable t)
    {
        var s = $"{t.SheetName} {t.SemesterTitle}".ToLowerInvariant();

        // варианты, которые обычно встречаются:
        // "доп", "дополнительная", "доп.нагрузка", "доп нагрузка"
        if (s.Contains("доп") || s.Contains("дополн"))
            return true;

        // если вдруг Excel листы на англ:
        if (s.Contains("extra") || s.Contains("additional"))
            return true;

        return false;
    }

    private IEnumerable<OpenXmlElement> BuildTeachingTables(IReadOnlyList<PlanTable> tables, bool isExtraLoad)
    {
        var list = new List<OpenXmlElement>();

        foreach (var t in tables.OrderBy(x => x.SheetName))
        {
            var dataRows = t.Rows
                .Where(r => !r.IsSummary)
                .OrderBy(r => r.RowOrder)
                .ToList();

            if (dataRows.Count == 0)
                continue;

            list.Add(TitleParagraph($"{t.SheetName} — {t.SemesterTitle}".Trim(' ', '—')));

            var table = NewTeachingTable(isExtraLoad);
            var headerRowHeight = isExtraLoad ? 300 : 420;
            var subHeaderRowHeight = isExtraLoad ? 420 : 680;
            var dataRowHeight = isExtraLoad ? 280 : 0;

            // 25 колонок (как TableGrid ниже)
            table.AppendChild(RowWithHeight(headerRowHeight,
                HeaderCell("№\nп.п.", vMergeRestart: true),                    // 1
                HeaderCell("Наименование\nдисциплины", vMergeRestart: true),   // 2
                HeaderCell("Факультет,\nспециальность,\nгруппа", vMergeRestart: true), // 3
                HeaderCell("Количество", gridSpan: 4),                         // 4-7
                HeaderCell("Количество часов по видам", gridSpan: 16),         // 8-23
                HeaderCell("Всего", vertical: true, vMergeRestart: true),      // 24
                HeaderCell("Примеч.", vertical: true, vMergeRestart: true)     // 25
            ));

            table.AppendChild(RowWithHeight(subHeaderRowHeight,
                HeaderCell("", vMergeContinue: true), // 1
                HeaderCell("", vMergeContinue: true), // 2
                HeaderCell("", vMergeContinue: true), // 3

                HeaderCell("курс", vertical: true),        // 4
                HeaderCell("потоков", vertical: true),     // 5
                HeaderCell("групп", vertical: true),       // 6
                HeaderCell("студентов", vertical: true),   // 7

                HeaderCell("лекц.", vertical: true),       // 8
                HeaderCell("практ.", vertical: true),      // 9
                HeaderCell("лаб.", vertical: true),        // 10
                HeaderCell("КСР", vertical: true),         // 11
                HeaderCell("КП", vertical: true),          // 12
                HeaderCell("КР", vertical: true),          // 13
                HeaderCell("контр.\nраб.", vertical: true),// 14
                HeaderCell("зач.", vertical: true),        // 15
                HeaderCell("диф.\nзач.", vertical: true),  // 16
                HeaderCell("экз.", vertical: true),        // 17
                HeaderCell("гос.\nэкз.", vertical: true),  // 18
                HeaderCell("ГЭК", vertical: true),         // 19
                HeaderCell("рук.\nВКР", vertical: true),    // 20
                HeaderCell("уч.\nпракт.", vertical: true),  // 21
                HeaderCell("произв.\nпракт.", vertical: true), // 22
                HeaderCell("предд.\nпракт.", vertical: true),  // 23

                HeaderCell("", vMergeContinue: true),      // 24
                HeaderCell("", vMergeContinue: true)       // 25
            ));

            foreach (var r in dataRows)
            {
                var row = dataRowHeight > 0
                    ? RowWithHeight(dataRowHeight,
                        CellText(r.Number?.ToString() ?? "", false, alignCenter: true),
                        CellText(r.DisciplineName ?? "", false),
                        CellText(r.FacultyGroup ?? "", false),

                        CellText(Val(r.Course), false, alignCenter: true),
                        CellText(Val(r.Streams), false, alignCenter: true),
                        CellText(Val(r.Groups), false, alignCenter: true),
                        CellText(Val(r.Students), false, alignCenter: true),

                        CellText(Val(r.Lek), false, alignCenter: true),
                        CellText(Val(r.Pr), false, alignCenter: true),
                        CellText(Val(r.Lab), false, alignCenter: true),
                        CellText(Val(r.Ksr), false, alignCenter: true),
                        CellText(Val(r.Kp), false, alignCenter: true),
                        CellText(Val(r.Kr), false, alignCenter: true),
                        CellText(Val(r.KontrolRab), false, alignCenter: true),
                        CellText(Val(r.Zach), false, alignCenter: true),
                        CellText(Val(r.DifZach), false, alignCenter: true),
                        CellText(Val(r.Exz), false, alignCenter: true),
                        CellText(Val(r.GosExz), false, alignCenter: true),
                        CellText(Val(r.Gek), false, alignCenter: true),
                        CellText(Val(r.RukVkr), false, alignCenter: true),

                        CellText(Val(r.UchPr), false, alignCenter: true),
                        CellText(Val(r.PrPr), false, alignCenter: true),
                        CellText(Val(r.PredPr), false, alignCenter: true),

                        CellText(Val(r.Total), false, alignCenter: true),
                        CellText(r.Note ?? "", false))
                    : Row(
                        CellText(r.Number?.ToString() ?? "", false, alignCenter: true),
                        CellText(r.DisciplineName ?? "", false),
                        CellText(r.FacultyGroup ?? "", false),

                        CellText(Val(r.Course), false, alignCenter: true),
                        CellText(Val(r.Streams), false, alignCenter: true),
                        CellText(Val(r.Groups), false, alignCenter: true),
                        CellText(Val(r.Students), false, alignCenter: true),

                        CellText(Val(r.Lek), false, alignCenter: true),
                        CellText(Val(r.Pr), false, alignCenter: true),
                        CellText(Val(r.Lab), false, alignCenter: true),
                        CellText(Val(r.Ksr), false, alignCenter: true),
                        CellText(Val(r.Kp), false, alignCenter: true),
                        CellText(Val(r.Kr), false, alignCenter: true),
                        CellText(Val(r.KontrolRab), false, alignCenter: true),
                        CellText(Val(r.Zach), false, alignCenter: true),
                        CellText(Val(r.DifZach), false, alignCenter: true),
                        CellText(Val(r.Exz), false, alignCenter: true),
                        CellText(Val(r.GosExz), false, alignCenter: true),
                        CellText(Val(r.Gek), false, alignCenter: true),
                        CellText(Val(r.RukVkr), false, alignCenter: true),

                        CellText(Val(r.UchPr), false, alignCenter: true),
                        CellText(Val(r.PrPr), false, alignCenter: true),
                        CellText(Val(r.PredPr), false, alignCenter: true),

                        CellText(Val(r.Total), false, alignCenter: true),
                        CellText(r.Note ?? "", false));

                table.AppendChild(row);
            }

            // Итоги снизу: без объединений (чтобы не было огромных ячеек)
            AppendTotalsRows(table, t);

            list.Add(table);
            list.Add(Spacer());
        }

        if (list.Count == 0)
            list.Add(new Paragraph(new Run(new Text(" "))));

        return list;
    }

    private IEnumerable<OpenXmlElement> BuildExtraLoadBlock(IReadOnlyList<PlanTable> tables)
    {
        var list = new List<OpenXmlElement>
        {
            SectionBreakLandscape()
        };

        list.AddRange(BuildTeachingTables(tables, isExtraLoad: true));
        list.Add(SectionBreakPortrait());
        return list;
    }

    private void AppendTotalsRows(Table table, PlanTable t)
    {
        // 25 колонок:
        // 1 пусто, 2 "Итого...", 3 "Поручено/Выполнено", 4..23 пусто, 24 Total, 25 Note
        void Add(string label, string kind, int? total)
        {
            table.AppendChild(Row(
                CellText("", true, alignCenter: true), // 1
                CellText(label, true),                 // 2
                CellText(kind, true),                  // 3

                // 4..23 пусто (20 колонок)
                CellText("", false), // 4
                CellText("", false), // 5
                CellText("", false), // 6
                CellText("", false), // 7
                CellText("", false), // 8
                CellText("", false), // 9
                CellText("", false), // 10
                CellText("", false), // 11
                CellText("", false), // 12
                CellText("", false), // 13
                CellText("", false), // 14
                CellText("", false), // 15
                CellText("", false), // 16
                CellText("", false), // 17
                CellText("", false), // 18
                CellText("", false), // 19
                CellText("", false), // 20
                CellText("", false), // 21
                CellText("", false), // 22
                CellText("", false), // 23

                CellText(total?.ToString() ?? "", true, alignCenter: true), // 24
                CellText("", false)                                         // 25
            ));
        }

        if (t.Sem1Plan.HasValue || t.Sem1Fact.HasValue)
        {
            Add(string.Empty, "Поручено", t.Sem1Plan);
            Add(string.Empty, "Выполнено", t.Sem1Fact);
        }

        if (t.Sem2Plan.HasValue || t.Sem2Fact.HasValue)
        {
            Add(string.Empty, "Поручено", t.Sem2Plan);
            Add(string.Empty, "Выполнено", t.Sem2Fact);
        }

        if (t.YearPlan.HasValue || t.YearFact.HasValue)
        {
            Add(string.Empty, "Поручено", t.YearPlan);
            Add(string.Empty, "Выполнено", t.YearFact);
        }
    }

    private OpenXmlElement BuildSummaryTable(IReadOnlyList<SummaryRow> rows)
    {
        var table = NewSummaryTable();

        table.AppendChild(RowWithHeight(420,
            HeaderCell("№\nп.п.", vMergeRestart: true),
            HeaderCell("Виды работ\n(в часах)", vMergeRestart: true),
            HeaderCell("1 семестр", gridSpan: 2),
            HeaderCell("2 семестр", gridSpan: 2),
            HeaderCell("На год", gridSpan: 2)
        ));

        table.AppendChild(RowWithHeight(360,
            HeaderCell("", vMergeContinue: true),
            HeaderCell("", vMergeContinue: true),
            HeaderCell("план"),
            HeaderCell("факт"),
            HeaderCell("план"),
            HeaderCell("факт"),
            HeaderCell("план"),
            HeaderCell("факт")
        ));

        foreach (var r in rows.OrderBy(r => r.RowOrder))
        {
            var isTotal = r.IsTotalRow;

            table.AppendChild(Row(
                CellText(r.Code ?? "", isTotal, alignCenter: true, fontSize: FontSizeBody),
                CellText(r.WorkName ?? "", isTotal, fontSize: FontSizeBody),
                CellText(Val(r.Sem1Plan), isTotal, alignCenter: true, fontSize: FontSizeBody),
                CellText(Val(r.Sem1Fact), isTotal, alignCenter: true, fontSize: FontSizeBody),
                CellText(Val(r.Sem2Plan), isTotal, alignCenter: true, fontSize: FontSizeBody),
                CellText(Val(r.Sem2Fact), isTotal, alignCenter: true, fontSize: FontSizeBody),
                CellText(Val(r.YearPlan), isTotal, alignCenter: true, fontSize: FontSizeBody),
                CellText(Val(r.YearFact), isTotal, alignCenter: true, fontSize: FontSizeBody)
            ));
        }

        return table;
    }

    private static string Val(int? x) => x?.ToString() ?? "";

    private Paragraph TitleParagraph(string text)
        => new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines { Before = "60", After = "120" }),
            new Run(
                new RunProperties(new Bold(), new FontSize { Val = FontSizeTitle }),
                new Text(text)));

    private static Paragraph Spacer() => new Paragraph(new Run(new Text(" ")));

    private static Paragraph SectionBreakLandscape()
    {
        var pageSize = new PageSize
        {
            Width = 16838,
            Height = 11906,
            Orient = PageOrientationValues.Landscape
        };

        var pageMargin = new PageMargin
        {
            Top = 720,
            Bottom = 720,
            Left = 720,
            Right = 720,
            Header = 720,
            Footer = 720,
            Gutter = 0
        };

        return new Paragraph(new ParagraphProperties(
            new SectionProperties(pageSize, pageMargin)));
    }

    private static Paragraph SectionBreakPortrait()
    {
        var pageSize = new PageSize
        {
            Width = 11906,
            Height = 16838,
            Orient = PageOrientationValues.Portrait
        };

        var pageMargin = new PageMargin
        {
            Top = 720,
            Bottom = 720,
            Left = 720,
            Right = 720,
            Header = 720,
            Footer = 720,
            Gutter = 0
        };

        return new Paragraph(new ParagraphProperties(
            new SectionProperties(pageSize, pageMargin)));
    }

    // -------------------- Table factories (FIXED layout + grids) --------------------

    private Table NewTeachingTable(bool isExtraLoad)
    {
        var table = NewTableBaseFixed();

        // 25 колонок (как Header/Rows/Totals)
        var grid = isExtraLoad
            ? new TableGrid(
                new GridColumn { Width = "520" },   // 1 №
                new GridColumn { Width = "4200" },  // 2 дисциплина
                new GridColumn { Width = "2600" },  // 3 группа

                new GridColumn { Width = "560" },   // 4 курс
                new GridColumn { Width = "600" },   // 5 потоков
                new GridColumn { Width = "600" },   // 6 групп
                new GridColumn { Width = "760" },   // 7 студентов

                new GridColumn { Width = "520" },   // 8 лек
                new GridColumn { Width = "520" },   // 9 пр
                new GridColumn { Width = "520" },   // 10 лаб
                new GridColumn { Width = "520" },   // 11 кср
                new GridColumn { Width = "500" },   // 12 кп
                new GridColumn { Width = "500" },   // 13 кр
                new GridColumn { Width = "600" },   // 14 контр раб
                new GridColumn { Width = "500" },   // 15 зач
                new GridColumn { Width = "600" },   // 16 диф зач
                new GridColumn { Width = "500" },   // 17 экз
                new GridColumn { Width = "600" },   // 18 госэкз
                new GridColumn { Width = "500" },   // 19 гэк
                new GridColumn { Width = "600" },   // 20 рук вкр
                new GridColumn { Width = "600" },   // 21 учпр
                new GridColumn { Width = "600" },   // 22 прпр
                new GridColumn { Width = "600" },   // 23 предпр

                new GridColumn { Width = "760" },   // 24 всего
                new GridColumn { Width = "1600" }   // 25 примеч.
            )
            : new TableGrid(
                new GridColumn { Width = "420" },   // 1 №
                new GridColumn { Width = "3400" },  // 2 дисциплина
                new GridColumn { Width = "2100" },  // 3 группа

                new GridColumn { Width = "480" },   // 4 курс
                new GridColumn { Width = "520" },   // 5 потоков
                new GridColumn { Width = "520" },   // 6 групп
                new GridColumn { Width = "650" },   // 7 студентов

                new GridColumn { Width = "430" },   // 8 лек
                new GridColumn { Width = "430" },   // 9 пр
                new GridColumn { Width = "430" },   // 10 лаб
                new GridColumn { Width = "430" },   // 11 кср
                new GridColumn { Width = "420" },   // 12 кп
                new GridColumn { Width = "420" },   // 13 кр
                new GridColumn { Width = "520" },   // 14 контр раб
                new GridColumn { Width = "420" },   // 15 зач
                new GridColumn { Width = "520" },   // 16 диф зач
                new GridColumn { Width = "420" },   // 17 экз
                new GridColumn { Width = "520" },   // 18 госэкз
                new GridColumn { Width = "420" },   // 19 гэк
                new GridColumn { Width = "520" },   // 20 рук вкр
                new GridColumn { Width = "520" },   // 21 учпр
                new GridColumn { Width = "520" },   // 22 прпр
                new GridColumn { Width = "520" },   // 23 предпр

                new GridColumn { Width = "620" },   // 24 всего
                new GridColumn { Width = "1200" }   // 25 примеч.
            );

        table.InsertAt(grid, 0);
        return table;
    }

    private Table NewSummaryTable()
    {
        var table = NewTableBaseFixed();

        var grid = new TableGrid(
            new GridColumn { Width = "520" },   // №
            new GridColumn { Width = "5200" },  // виды работ
            new GridColumn { Width = "900" },   // 1 сем план
            new GridColumn { Width = "900" },   // 1 сем факт
            new GridColumn { Width = "900" },   // 2 сем план
            new GridColumn { Width = "900" },   // 2 сем факт
            new GridColumn { Width = "900" },   // год план
            new GridColumn { Width = "900" }    // год факт
        );

        table.InsertAt(grid, 0);
        return table;
    }

    private static Table NewTableBaseFixed()
    {
        var table = new Table();

        var props = new TableProperties(
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }, // 100%
            new TableCellMarginDefault(
                new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 60, Type = TableWidthValues.Dxa },
                new TableCellRightMargin { Width = 60, Type = TableWidthValues.Dxa }
            ),
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 6 },
                new LeftBorder { Val = BorderValues.Single, Size = 6 },
                new BottomBorder { Val = BorderValues.Single, Size = 6 },
                new RightBorder { Val = BorderValues.Single, Size = 6 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
            )
        );

        table.AppendChild(props);
        return table;
    }

    // -------------------- element builders --------------------

    private static TableRow Row(params TableCell[] cells)
    {
        var r = new TableRow();
        foreach (var c in cells) r.AppendChild(c);
        return r;
    }

    private static TableRow RowWithHeight(int height, params TableCell[] cells)
    {
        var row = Row(cells);
        row.PrependChild(new TableRowProperties(
            new TableRowHeight { Val = (uint)height, HeightType = HeightRuleValues.AtLeast }));
        return row;
    }

    private static TableCell CellText(
        string text,
        bool bold,
        int gridSpan = 1,
        bool alignCenter = false,
        string fontSize = FontSizeBody)
    {
        var runProps = new RunProperties(new FontSize { Val = fontSize }, new NoProof());
        if (bold) runProps.Append(new Bold());

        var run = new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        var pPr = new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
        );

        if (alignCenter)
            pPr.Append(new Justification { Val = JustificationValues.Center });

        var para = new Paragraph(pPr, run);

        var tcProps = new TableCellProperties(
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
        );

        if (gridSpan > 1)
            tcProps.AppendChild(new GridSpan { Val = gridSpan });

        return new TableCell(tcProps, para);
    }

    private static TableCell HeaderCell(
        string text,
        int gridSpan = 1,
        bool vertical = false,
        bool vMergeRestart = false,
        bool vMergeContinue = false)
    {
        var tcProps = new TableCellProperties(
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
        );

        if (gridSpan > 1)
            tcProps.AppendChild(new GridSpan { Val = gridSpan });

        if (vMergeRestart)
            tcProps.AppendChild(new VerticalMerge { Val = MergedCellValues.Restart });
        else if (vMergeContinue)
            tcProps.AppendChild(new VerticalMerge { Val = MergedCellValues.Continue });

        if (vertical)
            tcProps.AppendChild(new TextDirection { Val = TextDirectionValues.BottomToTopLeftToRight });

        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
        );

        var runProps = new RunProperties(new FontSize { Val = FontSizeHeader }, new Bold(), new NoProof());
        var run = new Run(runProps);

        var lines = text.Split('\n', StringSplitOptions.None);
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                run.Append(new Break());
            run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
        }

        var para = new Paragraph(pPr, run);
        return new TableCell(tcProps, para);
    }
}
