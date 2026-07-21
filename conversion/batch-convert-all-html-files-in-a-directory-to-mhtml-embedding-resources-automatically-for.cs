using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output MHTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputMhtml");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        CreateSampleHtml(Path.Combine(inputFolder, "Sample1.html"), "<html><body><h1>Sample 1</h1><p>Hello World!</p></body></html>");
        CreateSampleHtml(Path.Combine(inputFolder, "Sample2.html"), "<html><body><h2>Sample 2</h2><img src=\"https://via.placeholder.com/150\" alt=\"Placeholder\"/></body></html>");

        // Process each HTML file in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Prepare save options for MHTML with embedded resources.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportFontResources = true,          // Embed font resources.
                ExportImagesAsBase64 = false,        // Keep images as separate MIME parts (default behavior).
                ExportCidUrlsForMhtmlResources = false // Use file name references (default).
            };

            // Determine output file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".mht";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as MHTML.
            doc.Save(outputPath, saveOptions);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"MHTML file was not created: {outputPath}");
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtml(string filePath, string htmlContent)
    {
        // Write deterministic HTML content to a file.
        File.WriteAllText(filePath, htmlContent);
    }
}
