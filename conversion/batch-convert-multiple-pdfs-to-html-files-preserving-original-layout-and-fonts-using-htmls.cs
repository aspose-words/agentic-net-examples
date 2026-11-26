using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files
        for (int i = 1; i <= 3; i++)
        {
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            CreateSamplePdf(pdfPath, $"Sample PDF Document {i}");
        }

        // Batch convert PDFs to HTML
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load PDF document (Aspose.Words can load PDF directly)
            Document pdfDoc = new Document(pdfFile);

            // Configure HTML save options to preserve layout and fonts
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportFontResources = true,   // Embed fonts
                ExportImagesAsBase64 = true,  // Embed images
                ExportPageMargins = true,     // Keep page margins
                PrettyFormat = true           // Indented HTML
            };

            // Determine output path
            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save as HTML
            pdfDoc.Save(htmlPath, htmlOptions);

            // Validate output
            if (!File.Exists(htmlPath) || new FileInfo(htmlPath).Length == 0)
            {
                throw new InvalidOperationException($"HTML conversion failed for '{pdfFile}'.");
            }
        }

        Console.WriteLine("Batch PDF to HTML conversion completed successfully.");
    }

    private static void CreateSamplePdf(string outputPath, string title)
    {
        // Create a simple Word document and save it as PDF
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);
        builder.Writeln("This is a sample paragraph to demonstrate layout preservation.");
        builder.InsertParagraph();
        builder.Writeln("Another line with some **bold** text.");

        // Save as PDF
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate creation
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException($"Failed to create sample PDF at '{outputPath}'.");
        }
    }
}
