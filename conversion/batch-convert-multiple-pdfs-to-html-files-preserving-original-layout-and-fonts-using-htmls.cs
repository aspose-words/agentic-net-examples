using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Directory.GetCurrentDirectory();

        // Prepare folders for input PDFs and output HTML files.
        string inputFolder = Path.Combine(baseDir, "InputPdfs");
        string outputFolder = Path.Combine(baseDir, "OutputHtml");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Number of sample PDF documents to create.
        const int documentCount = 3;

        // Create sample PDF files.
        for (int i = 1; i <= documentCount; i++)
        {
            // Build a simple Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample PDF document #{i}");
            builder.Writeln("This document is generated programmatically for batch conversion testing.");
            builder.Writeln($"Current date and time: {DateTime.Now}");

            // Save as PDF.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify PDF creation.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Batch convert each PDF to HTML while preserving layout and fonts.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Prepare folders for resources of this HTML conversion.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfFile);
            string htmlImagesFolder = Path.Combine(outputFolder, $"{fileNameWithoutExt}_Images");
            string htmlFontsFolder = Path.Combine(outputFolder, $"{fileNameWithoutExt}_Fonts");
            Directory.CreateDirectory(htmlImagesFolder);
            Directory.CreateDirectory(htmlFontsFolder);

            // Configure HtmlSaveOptions to export fonts and images to the folders above.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportFontResources = true,          // Preserve original fonts.
                ImagesFolder = htmlImagesFolder,     // Store linked images.
                FontsFolder = htmlFontsFolder,       // Store exported fonts.
                ExportTextInputFormFieldAsText = true // Example additional option.
            };

            // Save the HTML file.
            string htmlPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.html");
            pdfDoc.Save(htmlPath, htmlOptions);

            // Validate that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"HTML conversion failed for: {pdfFile}");
        }

        // Optional: indicate successful completion (no interactive output required).
        // The program will exit automatically.
    }
}
