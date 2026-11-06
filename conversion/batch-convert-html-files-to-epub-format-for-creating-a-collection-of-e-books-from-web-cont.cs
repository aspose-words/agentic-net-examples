using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputEpub");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample HTML files.
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample1.html"),
            "<html><body><h1>First Book</h1><p>This is the first e‑book.</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample2.html"),
            "<html><body><h1>Second Book</h1><p>This is the second e‑book.</p></body></html>");

        // Batch convert each HTML file to EPUB.
        foreach (string htmlPath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine the output EPUB file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string epubPath = Path.Combine(outputDir, fileNameWithoutExt + ".epub");

            // Save the document as EPUB.
            doc.Save(epubPath, SaveFormat.Epub);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
            {
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
            }
        }
    }

    private static void CreateSampleHtmlFile(string path, string htmlContent)
    {
        // Write deterministic HTML content to a file.
        File.WriteAllText(path, htmlContent);
    }
}
