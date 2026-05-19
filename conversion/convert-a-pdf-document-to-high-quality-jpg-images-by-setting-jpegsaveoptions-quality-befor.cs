using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");
        const string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a high‑quality JPEG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            // Configure JPEG save options with maximum quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 100,               // Highest quality (least compression).
                PageSet = new PageSet(i)          // Render the current page only.
            };

            string jpgPath = $"page_{i + 1}.jpg";
            pdfDoc.Save(jpgPath, jpegOptions);

            // Verify that the JPEG file was created.
            if (!File.Exists(jpgPath))
                throw new InvalidOperationException($"Failed to create JPEG image: {jpgPath}");
        }

        // Clean up the temporary PDF file if desired.
        // File.Delete(pdfPath);
    }
}
