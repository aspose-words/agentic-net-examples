using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF – this will be the source for image conversion.
        const string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(pdfPath);

        // Prepare output folder for PNG files.
        const string outputFolder = "OutputImages";
        Directory.CreateDirectory(outputFolder);

        // Export only even‑numbered pages (pages 2,4,…) as separate PNG images.
        // Page indices are zero‑based, so even‑numbered pages have odd indices.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Skip odd pages (index 0,2,4,… correspond to pages 1,3,5,…).
            if (pageIndex % 2 == 0) continue;

            // Render the current even page.
            pngOptions.PageSet = new PageSet(pageIndex);
            string pngPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            pdfDoc.Save(pngPath, pngOptions);

            // Validate that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Export of even‑numbered pages to PNG completed successfully.");
    }
}
