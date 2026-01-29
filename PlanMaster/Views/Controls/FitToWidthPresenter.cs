using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace PlanMaster.Views.Controls;

public sealed class FitToWidthPresenter : Decorator
{
    public static readonly StyledProperty<double> DesignWidthProperty =
        AvaloniaProperty.Register<FitToWidthPresenter, double>(nameof(DesignWidth), 1800);

    public double DesignWidth
    {
        get => GetValue(DesignWidthProperty);
        set => SetValue(DesignWidthProperty, value);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        var child = Child;
        if (child is null)
            return finalSize;

        var scale = DesignWidth <= 0 ? 1.0 : finalSize.Width / DesignWidth;
        if (scale > 1.0) scale = 1.0;
        if (scale < 0.35) scale = 0.35;

        child.RenderTransformOrigin = new RelativePoint(0, 0, RelativeUnit.Relative);
        child.RenderTransform = new ScaleTransform(scale, scale);

        var childSize = new Size(finalSize.Width / scale, finalSize.Height / scale);
        child.Arrange(new Rect(childSize));

        return finalSize;
    }
}