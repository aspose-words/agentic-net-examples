using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchPdfToHtml
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Directory.GetCurrentDirectory();

        // Prepare folders for input PDFs, output HTMLs, fonts and images.
        string inputFolder = Path.Combine(baseDir, "InputPdfs");
        string outputFolder = Path.Combine(baseDir, "OutputHtml");
        string fontsFolder = Path.Combine(baseDir, "ExportedFonts");
        string imagesFolder = Path.Combine(baseDir, "ExportedImages");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(fontsFolder);
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // Step 1: Create sample PDF files using Aspose.Words.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some sample content.
            builder.Writeln($"Sample PDF document #{i}");
            builder.Writeln("This document is generated programmatically for batch conversion testing.");
            builder.Writeln($"Current date and time: {DateTime.Now}");

            // Save as PDF in the input folder.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create input PDF: {pdfPath}");
        }

        // -----------------------------------------------------------------
        // Step 2: Batch convert each PDF to HTML while preserving layout and fonts.
        // -----------------------------------------------------------------
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure HtmlSaveOptions to export fonts and images.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,          // Preserve original fonts.
                FontsFolder = fontsFolder,           // Folder where font files will be written.
                ImagesFolder = imagesFolder,         // Folder where images will be written.
                // Optional: make the HTML more readable.
                PrettyFormat = true
            };

            // Determine the output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the document as HTML using the configured options.
            pdfDoc.Save(htmlPath, htmlOptions);

            // Validate that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"HTML conversion failed for: {pdfFile}");
        }

        // -----------------------------------------------------------------
        // Completion message (optional, not required for non‑interactive run).
        // -----------------------------------------------------------------
        // The program finishes here without waiting for user input.
    }
}
