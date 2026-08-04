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
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document that was just created.
        Document pdfDoc = new Document("input.pdf");

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Apply LZW compression.
            TiffCompression = TiffCompression.Lzw,
            // Set contrast to the maximum allowed value (0‑1 range). 
            // The original requirement of 10 is out of range and causes an exception.
            ImageContrast = 1.0f
        };

        // Save the PDF as a TIFF image using the specified options.
        pdfDoc.Save("output.tiff", tiffOptions);

        // Verify that the output file was created.
        if (!File.Exists("output.tiff"))
            throw new InvalidOperationException("The TIFF file was not created as expected.");
    }
}
