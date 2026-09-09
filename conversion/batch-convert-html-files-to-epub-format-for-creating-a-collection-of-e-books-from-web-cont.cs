using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output directories.
        string inputDir = "InputHtml";
        string outputDir = "OutputEpub";

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample HTML files.
        File.WriteAllText(Path.Combine(inputDir, "Sample1.html"),
            "<html><body><h1>First Document</h1><p>This is the first sample.</p></body></html>", Encoding.UTF8);
        File.WriteAllText(Path.Combine(inputDir, "Sample2.html"),
            "<html><body><h1>Second Document</h1><p>This is the second sample.</p></body></html>", Encoding.UTF8);

        // Process each HTML file in the input folder.
        foreach (string htmlFilePath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFilePath);

            // Configure save options for EPUB conversion.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                Encoding = Encoding.UTF8,
                // Optional: split the EPUB into parts by heading paragraphs.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                ExportDocumentProperties = true
            };

            // Determine the output EPUB file path.
            string epubFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".epub";
            string epubFilePath = Path.Combine(outputDir, epubFileName);

            // Save the document as EPUB.
            doc.Save(epubFilePath, saveOptions);

            // Verify that the EPUB file was created.
            if (!File.Exists(epubFilePath))
                throw new InvalidOperationException($"EPUB file was not created: {epubFilePath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
