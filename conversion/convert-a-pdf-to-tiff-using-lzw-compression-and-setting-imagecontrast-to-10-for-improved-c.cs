using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF created for conversion to TIFF.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample PDF.");

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Configure TIFF save options with LZW compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw
        };

        // Attempt to set image contrast; ignore if not supported.
        try
        {
            tiffOptions.ImageContrast = 10f;
        }
        catch (ArgumentOutOfRangeException)
        {
            // The property may not be supported for TIFF in some versions; continue without it.
        }

        // Save the PDF as a TIFF image.
        string tiffPath = "output.tiff";
        pdfDocument.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
            throw new InvalidOperationException("The TIFF conversion failed; no output file was created.");

        // Clean up temporary files (optional).
        File.Delete(pdfPath);
        // The TIFF file remains for inspection.
    }
}
