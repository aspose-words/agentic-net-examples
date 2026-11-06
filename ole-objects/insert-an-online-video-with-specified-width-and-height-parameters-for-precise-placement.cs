using System;
using System.IO;
using System.Diagnostics;

public class Program
{
    public static void Main()
    {
        // Video parameters
        string videoId = "dQw4w9WgXcQ"; // Example YouTube video ID
        int width = 560;
        int height = 315;

        // Build the embed URL
        string embedUrl = $"https://www.youtube.com/embed/{videoId}";

        // Create HTML content with the iframe
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"">
    <title>Embedded Video</title>
</head>
<body>
    <h2>Embedded YouTube Video</h2>
    <iframe width=""{width}"" height=""{height}"" src=""{embedUrl}""
            frameborder=""0"" allow=""accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture""
            allowfullscreen>
    </iframe>
</body>
</html>";

        // Write HTML to a temporary file
        string tempPath = Path.Combine(Path.GetTempPath(), $"EmbeddedVideo_{Guid.NewGuid()}.html");
        File.WriteAllText(tempPath, htmlContent);

        // Open the HTML file in the default browser
        OpenInDefaultBrowser(tempPath);
    }

    private static void OpenInDefaultBrowser(string filePath)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true // Let the OS decide how to open the file
            };
            Process.Start(psi);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to open the file automatically. Error: {ex.Message}");
            Console.WriteLine($"You can open the file manually at: {filePath}");
        }
    }
}
