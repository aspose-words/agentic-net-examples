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
            string htmlContent = $"<html><body><h1>Sample Document {i}</h1><p>This is a test HTML file.</p></body></html>";
            File.WriteAllText(Path.Combine(inputFolder, $"sample{i}.html"), htmlContent);
        }

        // Process each HTML file in the input folder.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Define a custom page size (e.g., 6 inches x 9 inches).
            // Aspose.Words uses points (1 inch = 72 points).
            const double inchesToPoints = 72.0;
            doc.FirstSection.PageSetup.PageWidth = 6 * inchesToPoints;   // 432 points
            doc.FirstSection.PageSetup.PageHeight = 9 * inchesToPoints;  // 648 points

            // Prepare PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF with the custom page size.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: indicate successful completion (no interactive prompts).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
