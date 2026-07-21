using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths for the intermediate PDF and final TIFF.
        const string pdfPath = "sample.pdf";
        const string tiffPath = "output.tiff";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample PDF content for conversion to TIFF.");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the generated PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure image save options:
        //    - Render to TIFF format.
        //    - Use LZW compression.
        //    - Increase image contrast (maximum allowed value is 1.0).
        // -----------------------------------------------------------------
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            ImageContrast = 1.0f // Maximum contrast within the valid range (0‑1).
        };

        // -----------------------------------------------------------------
        // 4. Save the PDF as a TIFF image using the configured options.
        // -----------------------------------------------------------------
        pdfDoc.Save(tiffPath, options);

        // -----------------------------------------------------------------
        // 5. Validate that the TIFF file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
        {
            throw new InvalidOperationException(
                "TIFF conversion failed: output file was not created or is empty.");
        }

        // Example completed successfully.
    }
}
