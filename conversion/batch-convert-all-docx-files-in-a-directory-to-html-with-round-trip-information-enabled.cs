using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            CreateSampleDocx(Path.Combine(inputFolder, "Sample1.docx"), "First sample document.");
            CreateSampleDocx(Path.Combine(inputFolder, "Sample2.docx"), "Second sample document with more text.\nLine two.\nLine three.");
        }

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Configure HTML save options to export round‑trip information.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportRoundtripInformation = true
            };

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(docxPath) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the document as HTML.
            doc.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Expected output HTML was not created: {htmlPath}");
        }
    }

    // Helper method to create a simple DOCX file with given content.
    private static void CreateSampleDocx(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath, SaveFormat.Docx);
    }
}
