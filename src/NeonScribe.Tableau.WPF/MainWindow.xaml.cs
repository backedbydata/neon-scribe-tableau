using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NeonScribe.Tableau.WPF.ViewModels;

namespace NeonScribe.Tableau.WPF;

public partial class MainWindow : Window
{
    private MainViewModel? ViewModel => DataContext as MainViewModel;

    public MainWindow()
    {
        InitializeComponent();
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        HandleFileDrop(e);
    }

    private void Window_DragOver(object sender, DragEventArgs e)
    {
        HandleDragOver(e);
    }

    private void DropZone_Drop(object sender, DragEventArgs e)
    {
        HandleFileDrop(e);
        ResetDropZoneAppearance();
    }

    private void DropZone_DragOver(object sender, DragEventArgs e)
    {
        HandleDragOver(e);
        HighlightDropZone();
    }

    private void DropZone_DragLeave(object sender, DragEventArgs e)
    {
        ResetDropZoneAppearance();
    }

    private void HandleFileDrop(DragEventArgs e)
    {
        if (ViewModel == null || ViewModel.IsProcessing)
            return;

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0)
            {
                var file = files[0];
                var extension = Path.GetExtension(file).ToLowerInvariant();

                if (extension == ".twb" || extension == ".twbx")
                {
                    ViewModel.SelectedFilePath = file;
                }
                else
                {
                    MessageBox.Show(
                        "Please drop a valid Tableau workbook file (.twb or .twbx).",
                        "Invalid File Type",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
        }

        e.Handled = true;
    }

    private void HandleDragOver(DragEventArgs e)
    {
        if (ViewModel == null || ViewModel.IsProcessing)
        {
            e.Effects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0)
            {
                var extension = Path.GetExtension(files[0]).ToLowerInvariant();
                e.Effects = (extension == ".twb" || extension == ".twbx")
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private void HighlightDropZone()
    {
        DropZone.BorderBrush = new SolidColorBrush(Color.FromRgb(33, 150, 243)); // Blue
        DropZone.BorderThickness = new Thickness(3);
    }

    private void ResetDropZoneAppearance()
    {
        DropZone.BorderThickness = new Thickness(2);
        DropZone.ClearValue(Border.BorderBrushProperty);
    }
}
