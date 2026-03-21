using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Returns a minimal 1x1 pixel PNG image.
    private static Stream CreatePlaceholderImage()
    {
        // PNG data for a 1x1 transparent pixel.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
            0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        return new MemoryStream(pngBytes);
    }

    // Generates a barcode image for DISPLAYBARCODE fields.
    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        // In a real scenario generate a barcode based on parameters.
        // Here we return a placeholder image.
        return CreatePlaceholderImage();
    }

    // Generates a barcode image for legacy BARCODE fields.
    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        // Return the same placeholder for simplicity.
        return CreatePlaceholderImage();
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field (Code128 barcode with sample text).
        builder.InsertField(@"DISPLAYBARCODE ""Code128"" ""1234567890""");

        // Assign the custom barcode generator to the document.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields so that the DISPLAYBARCODE field uses the generator.
        doc.UpdateFields();

        // Save the document as a PDF file.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
