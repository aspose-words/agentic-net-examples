using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document generated for conversion.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF document that was just created.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure JPEG save options:
        //    - High resolution (300 DPI)
        //    - Maximum JPEG quality (low compression)
        // -----------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            Resolution = 300,   // Sets both horizontal and vertical DPI.
            JpegQuality = 100   // 0 = worst quality, 100 = best quality.
        };

        // -----------------------------------------------------------------
        // 4. Save the PDF as a JPEG image.
        // -----------------------------------------------------------------
        const string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, jpegOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the JPEG file was created and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
        {
            throw new InvalidOperationException("JPEG conversion failed; output file is missing or empty.");
        }

        Console.WriteLine($"Conversion succeeded. JPEG saved to '{jpegPath}'.");
    }
}
