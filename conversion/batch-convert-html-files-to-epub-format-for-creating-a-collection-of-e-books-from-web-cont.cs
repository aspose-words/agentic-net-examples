using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToEpub
{
    public static void Main()
    {
        // Define folders for input HTML files and output EPUB files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputEpub");

        // Ensure the folders exist.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        Directory.CreateDirectory(inputFolder);

        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        CreateSampleHtml(Path.Combine(inputFolder, "Sample1.html"), "<html><body><h1>Chapter 1</h1><p>First chapter content.</p></body></html>");
        CreateSampleHtml(Path.Combine(inputFolder, "Sample2.html"), "<html><body><h1>Chapter 2</h1><p>Second chapter content.</p></body></html>");

        // Process each HTML file in the input folder.
        foreach (string htmlFilePath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFilePath);

            // Configure save options for EPUB output.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                SaveFormat = SaveFormat.Epub,
                Encoding = Encoding.UTF8,
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine the output EPUB file path.
            string epubFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".epub";
            string epubFilePath = Path.Combine(outputFolder, epubFileName);

            // Save the document as EPUB.
            doc.Save(epubFilePath, saveOptions);

            // Verify that the EPUB file was created.
            if (!File.Exists(epubFilePath))
                throw new InvalidOperationException($"EPUB file was not created: {epubFilePath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtml(string filePath, string htmlContent)
    {
        // Write deterministic HTML content to a file.
        File.WriteAllText(filePath, htmlContent, Encoding.UTF8);
    }
}
