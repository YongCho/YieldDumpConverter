using System.Windows.Input;

namespace YieldDumpConverter.Commands
{
    public static class CustomCommands
    {
        public static readonly RoutedUICommand OpenInExcelCommand = new RoutedUICommand(
            "Open In Excel",
            "OpenInExcelCommand",
            typeof(CustomCommands),
            new InputGestureCollection()
            {
                new KeyGesture(Key.E, ModifierKeys.Control | ModifierKeys.Shift)
            }
        );
    }
}
