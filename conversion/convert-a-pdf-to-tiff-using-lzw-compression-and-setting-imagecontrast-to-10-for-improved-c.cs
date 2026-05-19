using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion to TIFF.");

        // Save the document as PDF – this will be the input for conversion.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Apply LZW compression.
            TiffCompression = TiffCompression.Lzw,
            // Set contrast within the valid range (0‑1). Use a high value for clearer output.
            ImageContrast = 0.9f
        };

        // Save the PDF as a TIFF image using the specified options.
        const string tiffPath = "output.tiff";
        pdfDoc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("The expected TIFF output file was not created.");
    }
}
