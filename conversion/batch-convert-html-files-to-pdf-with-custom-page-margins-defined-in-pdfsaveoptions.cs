using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure clean start.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        CreateSampleHtml(Path.Combine(inputFolder, "sample1.html"), "<h1>Sample 1</h1><p>Hello from HTML 1.</p>");
        CreateSampleHtml(Path.Combine(inputFolder, "sample2.html"), "<h1>Sample 2</h1><p>Hello from HTML 2.</p>");

        // Batch conversion: each HTML file -> PDF with custom margins.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Set custom page margins (in points). 1 inch = 72 points.
            doc.FirstSection.PageSetup.TopMargin = 72;      // 1 inch
            doc.FirstSection.PageSetup.BottomMargin = 72;   // 1 inch
            doc.FirstSection.PageSetup.LeftMargin = 72;     // 1 inch
            doc.FirstSection.PageSetup.RightMargin = 72;    // 1 inch

            // Prepare PDF save options (no special settings required for margins).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, pdfOptions);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: indicate success (no console interaction required by the task).
    }

    private static void CreateSampleHtml(string filePath, string htmlContent)
    {
        // Write a simple HTML document to the specified file.
        string fullHtml = $"<!DOCTYPE html><html><head><meta charset=\"UTF-8\"></head><body>{htmlContent}</body></html>";
        File.WriteAllText(filePath, fullHtml);
    }
}
