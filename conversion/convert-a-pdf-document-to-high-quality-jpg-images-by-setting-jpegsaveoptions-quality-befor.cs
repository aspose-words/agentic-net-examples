using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple PDF document that will be used as the source.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This PDF will be converted to high‑quality JPEG images.");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // ---------------------------------------------------------------
        // 2. Load the PDF document that we just created.
        // ---------------------------------------------------------------
        Document pdfDocument = new Document(pdfPath);

        // ---------------------------------------------------------------
        // 3. Convert each page of the PDF to a JPEG image with high quality.
        // ---------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Configure image save options for JPEG format.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex),

                // Set JPEG quality to the maximum (100) for high‑quality output.
                JpegQuality = 100
            };

            string jpegPath = Path.Combine(outputDir, $"page_{pageIndex + 1}.jpg");
            pdfDocument.Save(jpegPath, jpegOptions);

            // Validate that the JPEG file was created and is not empty.
            if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
                throw new InvalidOperationException($"Failed to create JPEG for page {pageIndex + 1}.");
        }

        // Indicate successful completion.
        Console.WriteLine($"PDF conversion completed. Images are located in: {outputDir}");
    }
}
