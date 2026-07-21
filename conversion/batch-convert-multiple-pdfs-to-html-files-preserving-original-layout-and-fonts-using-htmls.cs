using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for the demonstration.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchConversion");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure clean folders.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a few sample PDF files that will be the source.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string pdfPath = Path.Combine(inputDir, $"Sample{i}.pdf");

            // Create a simple Word document and save it as PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of sample PDF #{i}.");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create source PDF: {pdfPath}");
        }

        // -----------------------------------------------------------------
        // Step 2: Batch convert each PDF in the input folder to HTML.
        // -----------------------------------------------------------------
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure HTML save options to preserve layout and fonts.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportFontResources = true, // Export fonts used in the document.
                // Store exported fonts in a subfolder next to the HTML file.
                FontsFolder = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(pdfFile) + "_fonts")
            };

            // Determine the output HTML file path.
            string htmlPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(pdfFile) + ".html");

            // Save the document as HTML using the configured options.
            pdfDoc.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created.
            if (!File.Exists(htmlPath))
                throw new InvalidOperationException($"HTML conversion failed for: {pdfFile}");
        }

        // -----------------------------------------------------------------
        // Completion message (optional, not required for the task).
        // -----------------------------------------------------------------
        Console.WriteLine("Batch PDF‑to‑HTML conversion completed successfully.");
    }
}
