using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfPagesToPng
{
    public static void Main()
    {
        // Create a sample multi‑page document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add five pages with simple text.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF – this will be the input for the conversion.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Prepare output directory for PNG files.
        const string outputFolder = "OutputImages";
        Directory.CreateDirectory(outputFolder);

        // Export only even‑numbered pages (page numbers 2,4,…) as separate PNG images.
        // In zero‑based indexing, even pages have odd indices.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            if (pageIndex % 2 == 1) // odd index => even page number
            {
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    // Use a PageSet that contains only the current page.
                    PageSet = new PageSet(pageIndex)
                };

                string pngPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
                pdfDoc.Save(pngPath, pngOptions);

                // Validate that the PNG file was written.
                if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                    throw new InvalidOperationException($"PNG for page {pageIndex + 1} was not created.");
            }
        }

        // Final validation: at least one PNG should exist.
        string[] generatedFiles = Directory.GetFiles(outputFolder, "*.png");
        if (generatedFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated.");

        // Example completed successfully.
        Console.WriteLine($"Exported {generatedFiles.Length} even‑page PNG file(s) to '{outputFolder}'.");
    }
}
