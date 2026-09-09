using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output HTML files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputHtml");
        string fontsFolder = Path.Combine(outputFolder, "Fonts");

        // Ensure a clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(fontsFolder);

        // Create sample PDF files.
        int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            // Create a blank document, add some content, and save as PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample PDF document #{i}");
            builder.Writeln("This document is generated for batch conversion testing.");
            builder.Writeln($"Current date and time: {DateTime.Now}");

            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Gather all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        List<string> convertedHtmlFiles = new List<string>();

        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure HtmlSaveOptions to export fonts and preserve layout.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportFontResources = true,
                FontsFolder = fontsFolder,
                // Do not embed images or CSS; keep them as external resources.
                ExportImagesAsBase64 = false
                // ExportEmbeddedCss and ExportEmbeddedImages are not members of HtmlSaveOptions.
            };

            // Determine output HTML file path.
            string htmlFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the document as HTML using the configured options.
            pdfDoc.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"Failed to create HTML file: {htmlPath}");

            convertedHtmlFiles.Add(htmlPath);
        }

        // Simple verification: ensure the expected number of HTML files were produced.
        if (convertedHtmlFiles.Count != pdfFiles.Length)
            throw new InvalidOperationException("The number of converted HTML files does not match the number of input PDFs.");

        // Optionally, output the result paths (commented out to avoid console interaction).
        // foreach (var html in convertedHtmlFiles)
        //     Console.WriteLine($"Converted: {html}");
    }
}
