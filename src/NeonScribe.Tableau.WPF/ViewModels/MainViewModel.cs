using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using NeonScribe.Tableau.WPF.Services;

namespace NeonScribe.Tableau.WPF.ViewModels;

public enum StatusState { Idle, Processing, Done, Error }

public partial class MainViewModel : ObservableObject
{
    private readonly DocumentationService _documentationService;

    [ObservableProperty]
    private string? _selectedFilePath;

    [ObservableProperty]
    private string _outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

    [ObservableProperty]
    private bool _isProcessing;

    [ObservableProperty]
    private string _statusMessage = string.Empty;

    [ObservableProperty]
    private StatusState _currentStatus = StatusState.Idle;

    public string? SelectedFileName => SelectedFilePath is null ? null : Path.GetFileName(SelectedFilePath);

    public bool CanGenerate => !string.IsNullOrEmpty(SelectedFilePath) && !IsProcessing;

    public MainViewModel(DocumentationService documentationService)
    {
        _documentationService = documentationService;
    }

    partial void OnSelectedFilePathChanged(string? value)
    {
        OnPropertyChanged(nameof(CanGenerate));
        OnPropertyChanged(nameof(SelectedFileName));
        GenerateDocumentationCommand.NotifyCanExecuteChanged();
        CurrentStatus = StatusState.Idle;
        StatusMessage = string.Empty;
    }

    partial void OnIsProcessingChanged(bool value)
    {
        OnPropertyChanged(nameof(CanGenerate));
        GenerateDocumentationCommand.NotifyCanExecuteChanged();
        BrowseFileCommand.NotifyCanExecuteChanged();
        BrowseOutputFolderCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanBrowse))]
    private void BrowseFile()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Tableau Workbook",
            Filter = "Tableau Files (*.twb;*.twbx)|*.twb;*.twbx|All Files (*.*)|*.*",
            DefaultExt = ".twb"
        };

        if (dialog.ShowDialog() == true)
        {
            SelectedFilePath = dialog.FileName;
        }
    }

    private bool CanBrowse() => !IsProcessing;

    [RelayCommand(CanExecute = nameof(CanBrowse))]
    private void BrowseOutputFolder()
    {
        var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog
        {
            Description = "Select Output Folder",
            UseDescriptionForTitle = true,
            SelectedPath = OutputDirectory
        };

        if (dialog.ShowDialog() == true)
        {
            OutputDirectory = dialog.SelectedPath;
        }
    }

    [RelayCommand(CanExecute = nameof(CanGenerate))]
    private async Task GenerateDocumentationAsync()
    {
        if (string.IsNullOrEmpty(SelectedFilePath))
            return;

        IsProcessing = true;
        CurrentStatus = StatusState.Processing;
        StatusMessage = "Generating...";

        try
        {
            if (!File.Exists(SelectedFilePath))
                throw new FileNotFoundException("The selected file does not exist.", SelectedFilePath);

            var extension = Path.GetExtension(SelectedFilePath).ToLowerInvariant();
            if (extension != ".twb" && extension != ".twbx")
                throw new InvalidOperationException("Please select a valid Tableau workbook file (.twb or .twbx).");

            var workbook = await _documentationService.ParseWorkbookAsync(SelectedFilePath);

            var inputFileName = Path.GetFileNameWithoutExtension(SelectedFilePath);
            var outputFilePath = Path.Combine(OutputDirectory, $"{inputFileName}-documentation.html");
            await _documentationService.GenerateHtmlAsync(workbook, outputFilePath);

            CurrentStatus = StatusState.Done;
            StatusMessage = "Opened in browser";

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = outputFilePath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            CurrentStatus = StatusState.Error;
            StatusMessage = ex.Message;

            MessageBox.Show(
                $"An error occurred while generating documentation:\n\n{ex.Message}",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsProcessing = false;
        }
    }
}
