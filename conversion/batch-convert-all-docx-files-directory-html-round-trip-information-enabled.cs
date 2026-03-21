using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocxToHtml
{
    static void Main()
    {
        // Folder containing the source DOCX files (relative to the executable).
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string sourceFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure both input and output directories exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Get all *.docx files in the source folder (non‑recursive).
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.TopDirectoryOnly);

        // Prepare the HtmlSaveOptions with round‑trip information enabled.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportRoundtripInformation = true // Preserve hidden elements, comments, etc.
        };

        foreach (string docxPath in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Build the output HTML file name (same base name, .html extension).
            string htmlFileName = Path.GetFileNameWithoutExtension(docxPath) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the document as HTML using the prepared options.
            doc.Save(htmlPath, htmlOptions);
        }

        Console.WriteLine("Conversion completed. {0} file(s) processed.", docxFiles.Length);
    }
}
