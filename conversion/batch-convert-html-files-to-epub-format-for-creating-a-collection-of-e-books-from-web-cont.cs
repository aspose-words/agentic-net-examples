using System;
using System.IO;
using Aspose.Words;

public class BatchHtmlToEpubConverter
{
    public static void Main()
    {
        // Define input and output directories
        string inputDir = "input_html";
        string outputDir = "output_epub";

        // Ensure clean start
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample HTML files
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample1.html"), "<html><body><h1>First Document</h1><p>This is the first sample.</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample2.html"), "<html><body><h1>Second Document</h1><p>This is the second sample.</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample3.html"), "<html><body><h1>Third Document</h1><p>This is the third sample.</p></body></html>");

        // Process each HTML file and convert to EPUB
        foreach (string htmlPath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load HTML document
            Document doc = new Document(htmlPath);

            // Determine output EPUB path
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string epubPath = Path.Combine(outputDir, fileNameWithoutExt + ".epub");

            // Save as EPUB
            doc.Save(epubPath, SaveFormat.Epub);

            // Validate output
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }

        // Optional: indicate success (no interactive output required)
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtmlFile(string path, string htmlContent)
    {
        File.WriteAllText(path, htmlContent);
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample HTML file: {path}");
    }
}
