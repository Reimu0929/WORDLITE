using System.Text;
using Microsoft.UI.Xaml;

namespace WordPadLite.WinUI;

public partial class App : Application
{
    private Window? _window;

    public App()
    {
        InitializeComponent();
        UnhandledException += OnUnhandledException;
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        _window = new MainWindow();
        _window.Activate();
    }

    private static void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        try
        {
            var path = Path.Combine(Path.GetTempPath(), "WordPadLite-crash.log");
            var text = new StringBuilder()
                .AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {e.Message}")
                .AppendLine(e.Exception?.ToString())
                .ToString();
            File.AppendAllText(path, text);
        }
        catch
        {
            // Never mask the original failure.
        }
    }
}
