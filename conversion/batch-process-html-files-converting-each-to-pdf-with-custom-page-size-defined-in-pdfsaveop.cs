using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "HtmlInput");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "PdfOutput");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        for (int i = 1; i <= 3; i++)
        {
            string htmlPath = Path.Combine(inputFolder, $"sample{i}.html");
            string htmlContent = $"<html><body><h1>Sample {i}</h1><p>This is HTML file number {i}.</p></body></html>";
            File.WriteAllText(htmlPath, htmlContent);
        }

        // Process each HTML file and convert to PDF with custom page size.
        foreach (string htmlFile in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load HTML document.
            Document doc = new Document(htmlFile);

            // Set a custom page size (width and height in points).
            // Example: 500 points wide x 700 points tall.
            doc.FirstSection.PageSetup.PageWidth = 500f;
            doc.FirstSection.PageSetup.PageHeight = 700f;

            // Prepare PDF save options (no special options needed beyond defaults).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, pdfOptions);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }

        // Final validation: ensure at least one PDF exists.
        string[] pdfFiles = Directory.GetFiles(outputFolder, "*.pdf");
        if (pdfFiles.Length == 0)
            throw new InvalidOperationException("No PDF files were generated.");

        // Example completed successfully.
        Console.WriteLine("Batch HTML to PDF conversion completed.");
    }
}
