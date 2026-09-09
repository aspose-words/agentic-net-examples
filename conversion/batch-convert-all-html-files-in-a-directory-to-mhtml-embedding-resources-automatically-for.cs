using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputHtml");
        string outputDir = Path.Combine(baseDir, "OutputMhtml");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample HTML files.
        for (int i = 1; i <= 3; i++)
        {
            string htmlPath = Path.Combine(inputDir, $"sample{i}.html");
            File.WriteAllText(htmlPath,
                $"<html><body><h1>Sample {i}</h1><p>This is a test HTML file.</p></body></html>");
        }

        // Convert each HTML file to MHTML with embedded resources.
        string[] htmlFiles = Directory.GetFiles(inputDir, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Configure save options for MHTML.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                // Use CID URLs for resources (optional, ensures better compatibility).
                ExportCidUrlsForMhtmlResources = true,
                // Export font resources so that the MHTML is fully self‑contained.
                ExportFontResources = true
            };

            // Determine the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".mht";
            string outputPath = Path.Combine(outputDir, outputFileName);

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
