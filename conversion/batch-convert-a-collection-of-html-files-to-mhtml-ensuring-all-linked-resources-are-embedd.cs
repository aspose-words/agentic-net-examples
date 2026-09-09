using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output MHTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputMhtml");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a simple PNG image (1x1 pixel) that will be referenced by the HTML files.
        string imagePath = Path.Combine(inputFolder, "sample.png");
        if (!File.Exists(imagePath))
        {
            // Base64 representation of a 1x1 transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ+Xc8AAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Create two sample HTML files that reference the PNG image.
        for (int i = 1; i <= 2; i++)
        {
            string htmlFileName = $"sample{i}.html";
            string htmlPath = Path.Combine(inputFolder, htmlFileName);

            if (!File.Exists(htmlPath))
            {
                string htmlContent = $@"
<!DOCTYPE html>
<html>
<head><title>Sample {i}</title></head>
<body>
    <h1>Sample Document {i}</h1>
    <p>This is a test HTML file.</p>
    <img src=""sample.png"" alt=""Sample Image"" />
</body>
</html>";
                File.WriteAllText(htmlPath, htmlContent, Encoding.UTF8);
            }
        }

        // Batch convert each HTML file in the input folder to MHTML.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Configure save options to embed all resources in the MHTML output.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true, // Use CID URLs for resources.
                ExportFontResources = true               // Ensure font resources are embedded if any.
            };

            // Determine the output MHTML file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as MHTML.
            doc.Save(outputPath, saveOptions);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create MHTML file: {outputPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
