using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BatchDocxToHtml
{
    static void Main()
    {
        // Folder containing the source DOCX files.
        string sourceFolder = @"C:\Docs\Input";

        // Folder where the converted HTML files will be saved.
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all *.docx files in the source folder (non‑recursive).
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.TopDirectoryOnly);

        // Prepare the HtmlSaveOptions with round‑trip information enabled.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportRoundtripInformation = true   // Preserve hidden elements, comments, etc.
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

        Console.WriteLine("Conversion completed. {0} files processed.", docxFiles.Length);
    }
}
