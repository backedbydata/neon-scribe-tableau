using System.IO.Compression;

namespace NeonScribe.Tableau.Core.Utilities;

public class TwbxExtractor
{
    public static string ExtractTwbFromTwbx(string twbxPath, string? outputDirectory = null)
    {
        if (!File.Exists(twbxPath))
        {
            throw new FileNotFoundException($"TWBX file not found: {twbxPath}");
        }

        if (!twbxPath.EndsWith(".twbx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("File must have .twbx extension", nameof(twbxPath));
        }

        // Use temp directory if no output directory specified
        outputDirectory ??= Path.Combine(Path.GetTempPath(), $"tableau-extract-{Guid.NewGuid()}");

        // Create output directory if it doesn't exist
        Directory.CreateDirectory(outputDirectory);

        // Extract the TWBX (it's a ZIP file)
        ZipFile.ExtractToDirectory(twbxPath, outputDirectory, overwriteFiles: true);

        // Find the TWB file in the extracted directory
        var twbFiles = Directory.GetFiles(outputDirectory, "*.twb", SearchOption.TopDirectoryOnly);

        if (twbFiles.Length == 0)
        {
            throw new InvalidOperationException($"No TWB file found in TWBX archive: {twbxPath}");
        }

        if (twbFiles.Length > 1)
        {
            throw new InvalidOperationException($"Multiple TWB files found in TWBX archive: {twbxPath}");
        }

        return twbFiles[0];
    }

    public static bool IsTwbxFile(string filePath)
    {
        return filePath.EndsWith(".twbx", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsTwbFile(string filePath)
    {
        return filePath.EndsWith(".twb", StringComparison.OrdinalIgnoreCase);
    }

    public static string GetTwbPath(string inputPath)
    {
        if (IsTwbFile(inputPath))
        {
            return inputPath;
        }
        else if (IsTwbxFile(inputPath))
        {
            return ExtractTwbFromTwbx(inputPath);
        }
        else
        {
            throw new ArgumentException("File must be either .twb or .twbx", nameof(inputPath));
        }
    }

    /// <summary>
    /// Extract images from TWBX file and return as dictionary of filename -> base64 data
    /// </summary>
    public static Dictionary<string, string> ExtractImagesFromTwbx(string twbxPath)
    {
        var images = new Dictionary<string, string>();

        if (!IsTwbxFile(twbxPath) || !File.Exists(twbxPath))
        {
            return images;
        }

        try
        {
            using var archive = ZipFile.OpenRead(twbxPath);

            // Look for common image formats in the archive
            var imageExtensions = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg" };

            foreach (var entry in archive.Entries)
            {
                var extension = Path.GetExtension(entry.Name).ToLowerInvariant();

                if (imageExtensions.Contains(extension))
                {
                    using var stream = entry.Open();
                    using var memoryStream = new MemoryStream();
                    stream.CopyTo(memoryStream);
                    var imageBytes = memoryStream.ToArray();
                    var base64 = Convert.ToBase64String(imageBytes);

                    // Store with the entry name as key
                    images[entry.Name] = base64;
                }
            }
        }
        catch
        {
            // If extraction fails, return empty dictionary
        }

        return images;
    }

    /// <summary>
    /// Get the directory where TWBX was extracted (for accessing images)
    /// </summary>
    public static string? GetExtractionDirectory(string twbPath)
    {
        // If the TWB is in a temp directory, that's the extraction directory
        if (twbPath.Contains(Path.GetTempPath()))
        {
            return Path.GetDirectoryName(twbPath);
        }

        return null;
    }
}
