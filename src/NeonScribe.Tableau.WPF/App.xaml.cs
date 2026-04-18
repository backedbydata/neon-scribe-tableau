using Microsoft.Extensions.DependencyInjection;
using System;
using System.Windows;
using NeonScribe.Tableau.WPF.Services;
using NeonScribe.Tableau.WPF.ViewModels;

namespace NeonScribe.Tableau.WPF;

public partial class App : Application
{
    private ServiceProvider? _serviceProvider;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Setup dependency injection
        var services = new ServiceCollection();
        ConfigureServices(services);
        _serviceProvider = services.BuildServiceProvider();

        // Show main window in the foreground
        var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
        mainWindow.Topmost = true;
        mainWindow.Activate();
        mainWindow.Topmost = false;
    }

    private void ConfigureServices(ServiceCollection services)
    {
        // Register services
        services.AddSingleton<DocumentationService>();

        // Register view models
        services.AddTransient<MainViewModel>();

        // Register views
        services.AddTransient<MainWindow>(provider =>
        {
            var viewModel = provider.GetRequiredService<MainViewModel>();
            var window = new MainWindow
            {
                DataContext = viewModel
            };
            return window;
        });
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _serviceProvider?.Dispose();
        base.OnExit(e);
    }
}
