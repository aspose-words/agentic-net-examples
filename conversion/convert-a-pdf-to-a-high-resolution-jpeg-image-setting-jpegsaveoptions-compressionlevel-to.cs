using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample Word document and save it as PDF (the source PDF).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion to JPEG.");
        const string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF document that was just created.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure image save options:
        //    - Use ImageSaveOptions with SaveFormat.Jpeg.
        //    - Set JpegQuality to 100 (lowest compression, highest quality).
        //    - Set a high DPI (Resolution) for better image detail.
        // -----------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            JpegQuality = 100,   // low compression = high quality
            Resolution = 300      // 300 DPI
        };

        // -----------------------------------------------------------------
        // 4. Save the first page of the PDF as a JPEG image.
        // -----------------------------------------------------------------
        const string jpegPath = "output.jpg";
        pdfDoc.Save(jpegPath, jpegOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the JPEG file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            throw new InvalidOperationException("The JPEG image was not created successfully.");
    }
}
