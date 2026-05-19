using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchPdfToHtmlConverter
{
    public static void Main()
    {
        // Base directory for the demo files.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ConversionDemo");
        string inputDir = Path.Combine(baseDir, "InputPdfs");
        string outputDir = Path.Combine(baseDir, "OutputHtml");
        string fontsDir = Path.Combine(outputDir, "Fonts");
        string imagesDir = Path.Combine(outputDir, "Images");

        // Ensure all required directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(fontsDir);
        Directory.CreateDirectory(imagesDir);

        // Create a few sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document pdfDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(pdfDoc);
            builder.Font.Name = "Arial";
            builder.Font.Size = 14;
            builder.Writeln($"Sample PDF document {i}");
            builder.Writeln("This PDF contains some sample text to demonstrate conversion.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page of the document.");
            string pdfPath = Path.Combine(inputDir, $"sample{i}.pdf");
            pdfDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to HTML while preserving layout and fonts.
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document doc = new Document(pdfFile);

            // Configure HTML save options to export font resources.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontsFolder = fontsDir,
                ImagesFolder = imagesDir
            };

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputDir, htmlFileName);

            // Save the document as HTML.
            doc.Save(htmlPath, htmlOptions);

            // Validate that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");
        }

        // Optional validation: ensure at least one font file was exported.
        string[] exportedFonts = Directory.GetFiles(fontsDir, "*.ttf");
        if (exportedFonts.Length == 0)
            throw new InvalidOperationException("No font resources were exported.");

        // Indicate successful completion.
        Console.WriteLine("Batch PDF to HTML conversion completed successfully.");
    }
}
