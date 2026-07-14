using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");

            // Create a blank document and add some content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample PDF number {i}.");
            builder.Writeln("The quick brown fox jumps over the lazy dog.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page of the document.");

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Batch convert each PDF to HTML while preserving layout and fonts.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Prepare HtmlSaveOptions to export fonts and preserve layout.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                // Store exported fonts in a dedicated folder per document.
                FontsFolder = Path.Combine(outputFolder, "fonts_" + Path.GetFileNameWithoutExtension(pdfFile))
            };

            // Ensure the fonts folder exists.
            Directory.CreateDirectory(htmlOptions.FontsFolder);

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the document as HTML using the specified options.
            pdfDoc.Save(htmlPath, htmlOptions);

            // Validate that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML: {htmlPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch PDF to HTML conversion completed successfully.");
    }
}
