using System;
using System.IO;
using System.Diagnostics;

public class Program
{
    public static void Main()
    {
        // Video parameters
        string videoUrl = "https://www.youtube.com/embed/dQw4w9WgXcQ";
        int width = 560;
        int height = 315;

        // Build HTML content with the embedded video
        string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Embedded Video</title>
</head>
<body>
    <h2>Embedded Online Video</h2>
    <iframe width=""{width}"" height=""{height}"" src=""{videoUrl}"" 
            frameborder=""0"" allow=""accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"" 
            allowfullscreen>
    </iframe>
</body>
</html>";

        // Write HTML to a temporary file
        string filePath = Path.Combine(Path.GetTempPath(), "EmbeddedVideo.html");
        File.WriteAllText(filePath, htmlContent);

        // Open the HTML file in the default browser
        var psi = new ProcessStartInfo
        {
            FileName = filePath,
            UseShellExecute = true
        };
        Process.Start(psi);
    }
}
