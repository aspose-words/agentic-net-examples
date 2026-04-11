using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directories for input HTML files and output MHTML files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputHtml");
        string outputDir = Path.Combine(baseDir, "OutputMhtml");

        // Ensure clean environment.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image (1x1 pixel) from a Base64 string.
        string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK8cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);
        string imageFileName = "sample.png";
        string imagePath = Path.Combine(inputDir, imageFileName);
        File.WriteAllBytes(imagePath, pngBytes);

        // Create two sample HTML files that reference the PNG image.
        for (int i = 1; i <= 2; i++)
        {
            string htmlContent = $@"<!DOCTYPE html>
<html>
<head><title>Sample {i}</title></head>
<body>
<h1>Document {i}</h1>
<p>This is a sample HTML file with an embedded image.</p>
<img src=""{imageFileName}"" alt=""Sample Image"" />
</body>
</html>";
            string htmlFileName = $"sample{i}.html";
            File.WriteAllText(Path.Combine(inputDir, htmlFileName), htmlContent, Encoding.UTF8);
        }

        // Process each HTML file in the input directory.
        foreach (string htmlFilePath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document. Resources (like images) are resolved relative to the input folder.
            Document doc = new Document(htmlFilePath);

            // Configure save options for MHTML. Resources will be embedded automatically.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs for resources – improves compatibility with some mail agents.
                ExportCidUrlsForMhtmlResources = true,
                // Ensure any linked fonts are also embedded (optional, but satisfies "all linked resources").
                ExportFontResources = true,
                // Embed images directly as Base64 within the MHTML (optional, also ensures embedding).
                ExportImagesAsBase64 = true
            };

            // Determine output file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".mht";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Save the document as MHTML.
            doc.Save(outputPath, saveOptions);

            // Validation: output file must exist and have non‑zero length.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException($"Failed to create MHTML file: {outputPath}");
            if (new FileInfo(outputPath).Length == 0)
                throw new InvalidOperationException($"MHTML file is empty: {outputPath}");
        }

        // Indicate successful batch conversion.
        Console.WriteLine($"Batch conversion completed. MHTML files are located in: {outputDir}");
    }
}
