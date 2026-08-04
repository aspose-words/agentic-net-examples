using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string jpegPath = "sample.jpg";

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);
        builder.Writeln("This is a sample PDF created with Aspose.Words.");
        builder.Writeln("It will be converted to a high‑resolution JPEG image.");
        pdfDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document loadedPdf = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Convert the first page of the PDF to a JPEG image.
        //    Use high resolution (300 DPI) and set JPEG quality to 100
        //    which corresponds to low compression (high quality).
        // -----------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            Resolution = 300f,   // High resolution.
            JpegQuality = 100    // Low compression, high quality.
        };

        loadedPdf.Save(jpegPath, jpegOptions);

        // Verify that the JPEG was created.
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("The JPEG image was not created successfully.");

        // Cleanup (optional): delete temporary files if desired.
        // File.Delete(pdfPath);
        // File.Delete(jpegPath);
    }
}
