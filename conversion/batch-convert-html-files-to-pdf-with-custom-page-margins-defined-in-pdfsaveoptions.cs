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
        CreateSampleHtml(Path.Combine(inputFolder, "Sample1.html"), "<h1>First Document</h1><p>Hello from HTML 1.</p>");
        CreateSampleHtml(Path.Combine(inputFolder, "Sample2.html"), "<h1>Second Document</h1><p>Hello from HTML 2.</p>");

        // Define custom margins (in points). 72 points = 1 inch.
        const double marginTop = 72;
        const double marginBottom = 72;
        const double marginLeft = 72;
        const double marginRight = 72;

        // Process each HTML file.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Apply custom margins to every section.
            foreach (Section section in doc.Sections)
            {
                section.PageSetup.TopMargin = marginTop;
                section.PageSetup.BottomMargin = marginBottom;
                section.PageSetup.LeftMargin = marginLeft;
                section.PageSetup.RightMargin = marginRight;
            }

            // Prepare PDF save options (optional customizations can be added here).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleHtml(string filePath, string htmlContent)
    {
        // Write a simple HTML document to the specified path.
        string fullHtml = $"<!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>Sample</title></head><body>{htmlContent}</body></html>";
        File.WriteAllText(filePath, fullHtml);
    }
}
