using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;

namespace EXOKit.Controls;

public sealed class PointerGrid : Grid
{
    private InputSystemCursor? _arrow;
    private InputSystemCursor? _hand;

    public PointerGrid()
    {
        Loaded += (_, _) =>
        {
            _arrow = InputSystemCursor.Create(InputSystemCursorShape.Arrow);
            _hand = InputSystemCursor.Create(InputSystemCursorShape.Hand);
            ProtectedCursor = _arrow;
        };
        Unloaded += (_, _) =>
        {
            ProtectedCursor = null;
            _arrow?.Dispose();
            _hand?.Dispose();
            _arrow = _hand = null;
        };
        AddHandler(PointerMovedEvent, new PointerEventHandler(UpdateCursor), true);
        AddHandler(PointerEnteredEvent, new PointerEventHandler(UpdateCursor), true);
        AddHandler(PointerReleasedEvent, new PointerEventHandler(UpdateCursor), true);
        PointerExited += (_, _) => ProtectedCursor = _arrow;
    }

    private void UpdateCursor(object sender, PointerRoutedEventArgs args)
    {
        var actionable = false;
        for (var element = args.OriginalSource as DependencyObject; element is UIElement; element = VisualTreeHelper.GetParent(element))
        {
            if (element is Control { IsEnabled: false }) break;
            if (element is TextBox or PasswordBox or RichEditBox or Thumb or ScrollBar
                || element is TextBlock { IsTextSelectionEnabled: true }) break;
            if (element is ButtonBase or ComboBox or ToggleSwitch or NavigationViewItem or PivotHeaderItem)
            {
                actionable = true;
                break;
            }
            if (element is ListViewItem item && ItemsControl.ItemsControlFromItemContainer(item) is ListView list)
            {
                actionable = list.SelectionMode != ListViewSelectionMode.None || list.IsItemClickEnabled;
                break;
            }
        }
        ProtectedCursor = actionable ? _hand : _arrow;
    }
}

public sealed class OutputResizeGrip : Grid
{
    private InputSystemCursor? _cursor;

    public OutputResizeGrip()
    {
        Loaded += (_, _) => ProtectedCursor = _cursor = InputSystemCursor.Create(InputSystemCursorShape.SizeNorthSouth);
        Unloaded += (_, _) =>
        {
            ProtectedCursor = null;
            _cursor?.Dispose();
            _cursor = null;
        };
    }
}

public sealed class PointerMenuFlyoutItem : MenuFlyoutItem
{
    private InputSystemCursor? _cursor;
    private InputSystemCursor? _arrow;

    public PointerMenuFlyoutItem()
    {
        Loaded += (_, _) =>
        {
            _cursor = InputSystemCursor.Create(InputSystemCursorShape.Hand);
            _arrow = InputSystemCursor.Create(InputSystemCursorShape.Arrow);
            ProtectedCursor = IsEnabled ? _cursor : _arrow;
        };
        IsEnabledChanged += (_, _) => ProtectedCursor = IsEnabled ? _cursor : _arrow;
        Unloaded += (_, _) =>
        {
            ProtectedCursor = null;
            _cursor?.Dispose();
            _arrow?.Dispose();
            _cursor = _arrow = null;
        };
    }
}