using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files if they do not already exist.
        for (int i = 1; i <= 3; i++)
        {
            string docxPath = Path.Combine(inputDir, $"Sample{i}.docx");
            if (!File.Exists(docxPath))
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document {i}.");
                sampleDoc.Save(docxPath, SaveFormat.Docx);
            }
        }

        // Convert each DOCX file in the input folder to HTML with round‑trip information.
        string[] docxFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxFile);

            // Configure HTML save options to include round‑trip information.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportRoundtripInformation = true
            };

            // Determine the output HTML file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxFile);
            string htmlPath = Path.Combine(outputDir, $"{fileNameWithoutExt}.html");

            // Save the document as HTML.
            doc.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");
        }
    }
}
