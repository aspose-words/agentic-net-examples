using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "InputHtml";
        string outputDir = "OutputEpub";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample HTML files.
        string htmlFile1 = Path.Combine(inputDir, "Sample1.html");
        string htmlContent1 = "<html><body><h1>First Document</h1><p>This is the first sample HTML file.</p></body></html>";
        File.WriteAllText(htmlFile1, htmlContent1, Encoding.UTF8);

        string htmlFile2 = Path.Combine(inputDir, "Sample2.html");
        string htmlContent2 = "<html><body><h1>Second Document</h1><p>This is the second sample HTML file.</p></body></html>";
        File.WriteAllText(htmlFile2, htmlContent2, Encoding.UTF8);

        // Process each HTML file in the input directory.
        string[] htmlFiles = Directory.GetFiles(inputDir, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Configure save options for EPUB.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Epub;
            saveOptions.Encoding = Encoding.UTF8;
            // Optional: split the EPUB into parts at heading paragraphs.
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
            // Export document properties (optional).
            saveOptions.ExportDocumentProperties = true;

            // Determine the output EPUB file path.
            string epubFileName = Path.ChangeExtension(Path.GetFileName(htmlPath), ".epub");
            string epubPath = Path.Combine(outputDir, epubFileName);

            // Save the document as EPUB.
            doc.Save(epubPath, saveOptions);

            // Verify that the EPUB file was created.
            if (!File.Exists(epubPath))
                throw new InvalidOperationException($"EPUB file was not created: {epubPath}");
        }

        // Indicate completion.
        Console.WriteLine("Batch conversion of HTML files to EPUB completed successfully.");
    }
}
