using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document with two pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the first page of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page of the sample PDF.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document we just created.
        Document pdfDoc = new Document(pdfPath);

        // Prepare image save options for high‑quality JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 100,                 // Maximum quality.
            UseHighQualityRendering = true,    // Enable high‑quality rendering algorithms.
            UseAntiAliasing = true             // Enable anti‑aliasing for smoother edges.
        };

        // Export each page of the PDF as a separate JPEG image.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            jpegOptions.PageSet = new PageSet(pageIndex);
            string outputFile = $"page_{pageIndex + 1}.jpg";
            pdfDoc.Save(outputFile, jpegOptions);

            // Verify that the image file was created.
            if (!File.Exists(outputFile))
                throw new InvalidOperationException($"Failed to create JPEG image: {outputFile}");
        }

        // Verify that the source PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF document.");
    }
}
