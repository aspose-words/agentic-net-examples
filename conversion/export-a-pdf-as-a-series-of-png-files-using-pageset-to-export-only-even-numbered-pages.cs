using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportPdfEvenPagesToPng
{
    public static void Main()
    {
        // Create a sample document with several pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 6; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 6)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF – this will be the source PDF.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(pdfPath);

        // Export only the even‑numbered pages (pages 2,4,6, ...) as separate PNG files.
        // In zero‑based indexing, even‑numbered pages have odd indices.
        for (int pageIndex = 1; pageIndex < pdfDoc.PageCount; pageIndex += 2)
        {
            // Configure image save options for PNG and select a single page.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex) // Render the current even page.
            };

            string outFile = $"Page_{pageIndex + 1}.png";
            pdfDoc.Save(outFile, pngOptions);

            // Verify that the PNG file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create PNG file: {outFile}");
        }

        // Verify that at least one PNG file exists.
        string[] pngFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "Page_*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG files were generated.");

        // Example completed successfully.
        Console.WriteLine("Exported even pages to PNG files:");
        foreach (string file in pngFiles)
            Console.WriteLine(file);
    }
}
