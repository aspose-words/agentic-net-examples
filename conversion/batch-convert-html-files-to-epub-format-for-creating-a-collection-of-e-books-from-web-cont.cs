using System;
using System.IO;
using Aspose.Words;

public class BatchHtmlToEpubConverter
{
    public static void Main()
    {
        // Define folders for input HTML files and output EPUB files.
        string inputFolder = "HtmlInputs";
        string outputFolder = "EpubOutputs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files in the input folder.
        CreateSampleHtmlFile(Path.Combine(inputFolder, "sample1.html"), "<html><body><h1>Sample 1</h1><p>This is the first sample.</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputFolder, "sample2.html"), "<html><body><h1>Sample 2</h1><p>This is the second sample.</p></body></html>");

        // Get all HTML files in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html", SearchOption.TopDirectoryOnly);

        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Determine the output EPUB file path.
            string outputFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".epub";
            string epubPath = Path.Combine(outputFolder, outputFileName);

            // Save the document as EPUB.
            doc.Save(epubPath, SaveFormat.Epub);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
            {
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
            }
        }

        // Optional: indicate successful conversion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtmlFile(string path, string htmlContent)
    {
        // Write deterministic HTML content to the specified file.
        File.WriteAllText(path, htmlContent);
    }
}
