using System.Reflection;
using System.Windows;

namespace Uninstaller.App.Views;

public partial class AboutWindow : Window
{
    public AboutWindow()
    {
        InitializeComponent();
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        VersionText.Text = $"Version {version?.ToString(3) ?? "1.0.0"}";
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
}
