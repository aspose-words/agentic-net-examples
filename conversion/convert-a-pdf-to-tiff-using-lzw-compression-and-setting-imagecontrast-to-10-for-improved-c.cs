using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion to TIFF.");
        const string pdfPath = "input.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure TIFF save options: LZW compression and increased contrast.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            // ImageContrast must be in the range 0..1. 1.0 gives maximum contrast.
            ImageContrast = 1.0f
        };

        // Save the PDF as a TIFF image.
        const string tiffPath = "output.tiff";
        pdfDoc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("The TIFF output file was not created.");

        // Optional cleanup of the intermediate PDF.
        // File.Delete(pdfPath);
    }
}
