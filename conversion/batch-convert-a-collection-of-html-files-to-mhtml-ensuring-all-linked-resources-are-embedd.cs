using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToMhtml
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputHtml");
        string outputDir = Path.Combine(baseDir, "OutputMhtml");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample images (1x1 pixel PNG)
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=");
        File.WriteAllBytes(Path.Combine(inputDir, "image1.png"), pngBytes);
        File.WriteAllBytes(Path.Combine(inputDir, "image2.png"), pngBytes);

        // Create sample HTML files that reference the images
        string html1 = @"<html><body><h1>Sample 1</h1><p>Image below:</p><img src=""image1.png"" alt=""Image1""></body></html>";
        string html2 = @"<html><body><h1>Sample 2</h1><p>Another image:</p><img src=""image2.png"" alt=""Image2""></body></html>";
        File.WriteAllText(Path.Combine(inputDir, "sample1.html"), html1);
        File.WriteAllText(Path.Combine(inputDir, "sample2.html"), html2);

        // Batch conversion: each HTML file -> MHTML with embedded resources
        foreach (string htmlPath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document
            Document doc = new Document(htmlPath);

            // Configure save options for MHTML
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportCidUrlsForMhtmlResources = true, // Use CID URLs for resources
                ExportFontResources = true               // Ensure fonts are embedded if any
            };

            // Determine output file name
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string mhtmlPath = Path.Combine(outputDir, fileNameWithoutExt + ".mht");

            // Save as MHTML
            doc.Save(mhtmlPath, saveOptions);

            // Validate that the MHTML file was created and contains data
            if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
                throw new InvalidOperationException($"MHTML conversion failed for '{htmlPath}'.");
        }

        // All conversions completed successfully
        Console.WriteLine("Batch conversion completed. MHTML files are located in: " + outputDir);
    }
}
