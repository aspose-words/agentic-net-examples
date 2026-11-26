using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDFs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure clean state.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        for (int i = 1; i <= 3; i++)
        {
            string htmlFilePath = Path.Combine(inputFolder, $"sample{i}.html");
            string htmlContent = $"<html><body><h1>Sample {i}</h1><p>This is HTML file number {i}.</p></body></html>";
            File.WriteAllText(htmlFilePath, htmlContent);
        }

        // Process each HTML file and convert to PDF with a custom page size.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Set a custom page size (595x842 points ≈ A4) for the first section.
            if (doc.FirstSection?.PageSetup != null)
            {
                doc.FirstSection.PageSetup.PageWidth = 595f;   // Width in points.
                doc.FirstSection.PageSetup.PageHeight = 842f; // Height in points.
            }

            // Configure PDF save options (no special options needed for custom size).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(htmlPath);
            string pdfPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.pdf");

            // Save the document as PDF using the options.
            doc.Save(pdfPath, pdfOptions);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }

        // Final validation: ensure the number of PDFs matches the number of HTML files.
        string[] pdfFiles = Directory.GetFiles(outputFolder, "*.pdf");
        if (pdfFiles.Length != htmlFiles.Length)
            throw new InvalidOperationException("The number of generated PDF files does not match the number of input HTML files.");

        // Example completed successfully.
        Console.WriteLine("Batch HTML to PDF conversion completed.");
    }
}
