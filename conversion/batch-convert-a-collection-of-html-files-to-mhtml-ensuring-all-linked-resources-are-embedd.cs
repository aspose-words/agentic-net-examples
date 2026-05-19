using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToMhtml
{
    public static void Main()
    {
        // Prepare folders.
        string inputFolder = "InputHtml";
        string outputFolder = "OutputMhtml";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a simple PNG image (1x1 pixel) from a Base64 string.
        string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X5eUAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);
        string imagePath = Path.Combine(inputFolder, "sample.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a simple CSS file.
        string cssContent = "body { font-family: Arial; } .highlight { color: red; }";
        string cssPath = Path.Combine(inputFolder, "style.css");
        File.WriteAllText(cssPath, cssContent, Encoding.UTF8);

        // Create two sample HTML files that reference the image and CSS.
        for (int i = 1; i <= 2; i++)
        {
            string html = $@"
<!DOCTYPE html>
<html>
<head>
    <title>Sample {i}</title>
    <link rel=""stylesheet"" type=""text/css"" href=""style.css"">
</head>
<body>
    <h1 class=""highlight"">Hello World {i}</h1>
    <p>This is a sample HTML file.</p>
    <img src=""sample.png"" alt=""Sample Image"">
</body>
</html>";
            string htmlPath = Path.Combine(inputFolder, $"sample{i}.html");
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
        }

        // Batch convert each HTML file to MHTML with all resources embedded.
        foreach (string htmlFile in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Configure save options for MHTML.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Ensure that fonts, images, CSS, etc., are embedded.
                ExportFontResources = true,
                ExportCidUrlsForMhtmlResources = true,
                CssStyleSheetType = CssStyleSheetType.External,
                PrettyFormat = true
            };

            // Determine output file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save as MHTML.
            doc.Save(outputPath, saveOptions);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create MHTML file: {outputPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
