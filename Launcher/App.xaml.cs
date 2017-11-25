using System.Windows;
using System.Windows.Controls;

namespace Launcher
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// Add events on startup for textbox
        /// </summary>
        protected override void OnStartup(StartupEventArgs e)
        {
            //Tab into textbox
            EventManager.RegisterClassHandler(typeof(TextBox),
                TextBox.GotFocusEvent,
                new RoutedEventHandler(TextBox_GotFocus));

            base.OnStartup(e);
        }

        /// <summary>
        /// Select all the text on text box GotFocus
        /// </summary>
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as TextBox).SelectAll();
        }
    }
}
