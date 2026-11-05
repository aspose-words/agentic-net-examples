using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "input");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files in the input folder.
        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some content to demonstrate layout preservation.
            builder.Writeln($"Sample PDF Document {i}");
            builder.Writeln("This is a paragraph with some sample text to illustrate conversion.");
            builder.InsertParagraph();
            builder.Writeln("Another line of text.");

            // Save the document as PDF.
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file '{pdfPath}' was not created.");
        }

        // Batch convert each PDF to HTML while preserving layout and fonts.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure HTML save options to preserve layout and fonts.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportPageMargins = true,
                ExportPageSetup = true,
                ExportFontResources = true,
                ExportImagesAsBase64 = true
                // ExportHeadersFootersMode omitted because the enum may not be available in all versions.
            };

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the PDF as HTML.
            pdfDoc.Save(htmlPath, htmlOptions);

            // Validate that the HTML file was created and contains data.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"HTML file '{htmlPath}' was not created.");

            FileInfo info = new FileInfo(htmlPath);
            if (info.Length == 0)
                throw new InvalidOperationException($"HTML file '{htmlPath}' is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch PDF to HTML conversion completed successfully.");
    }
}
