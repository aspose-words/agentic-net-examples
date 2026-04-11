using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample HTML files.
        for (int i = 1; i <= 3; i++)
        {
            string htmlContent = $"<html><body><p>Hello from sample HTML file {i}.</p></body></html>";
            string htmlPath = Path.Combine(inputFolder, $"sample{i}.html");
            File.WriteAllText(htmlPath, htmlContent);
        }

        // Process each HTML file in the input folder.
        foreach (string htmlFilePath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFilePath);

            // Define a custom page size (points). Here we use 500 x 800 points.
            doc.FirstSection.PageSetup.PaperSize = PaperSize.Custom;
            doc.FirstSection.PageSetup.PageWidth = 500;   // Width in points.
            doc.FirstSection.PageSetup.PageHeight = 800;  // Height in points.

            // Create PdfSaveOptions – can be used to set additional PDF-specific settings.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".pdf";
            string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the custom options.
            doc.Save(pdfFilePath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfFilePath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfFilePath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
