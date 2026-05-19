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
        CreateSampleHtml(Path.Combine(inputFolder, "Sample1.html"), "<html><body><h1>First Document</h1><p>Hello, world!</p></body></html>");
        CreateSampleHtml(Path.Combine(inputFolder, "Sample2.html"), "<html><body><h2>Second Document</h2><p>Another paragraph.</p></body></html>");

        // Process each HTML file in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Prepare save options for EPUB format.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                Encoding = Encoding.UTF8,
                ExportDocumentProperties = true,
                DocumentSplitCriteria = DocumentSplitCriteria.None
            };

            // Determine the output EPUB file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string epubPath = Path.Combine(outputFolder, fileNameWithoutExt + ".epub");

            // Save the document as EPUB.
            doc.Save(epubPath, saveOptions);

            // Validate that the EPUB file was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtml(string path, string htmlContent)
    {
        // Write deterministic HTML content to a file.
        File.WriteAllText(path, htmlContent, Encoding.UTF8);
    }
}
