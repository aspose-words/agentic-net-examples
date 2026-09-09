using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            string docxPath = Path.Combine(inputFolder, $"Sample{i}.docx");

            // Create a blank document and add some text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");

            // Save the document as DOCX.
            doc.Save(docxPath, SaveFormat.Docx);

            // Verify the DOCX was created.
            if (!File.Exists(docxPath))
                throw new InvalidOperationException($"Failed to create sample DOCX: {docxPath}");
        }

        // Batch convert each DOCX file to HTML with round‑trip information enabled.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            // Load the DOCX file.
            Document document = new Document(docxFile);

            // Configure HTML save options to export round‑trip information.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportRoundtripInformation = true
            };

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(docxFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the document as HTML.
            document.Save(htmlPath, htmlOptions);

            // Verify the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");
        }

        // Indicate successful conversion.
        Console.WriteLine($"Converted {docxFiles.Length} DOCX file(s) to HTML in '{outputFolder}'.");
    }
}
