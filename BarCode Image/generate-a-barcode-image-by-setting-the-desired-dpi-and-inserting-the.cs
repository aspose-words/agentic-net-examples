using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Assign a custom barcode generator. In a real scenario you would implement
        // the generator to produce actual barcode images. Here we use a stub.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Define barcode parameters. The ScalingFactor can be used to influence the
        // size (and effectively the DPI) of the generated image.
        BarcodeParameters barcodeParameters = new BarcodeParameters
        {
            BarcodeType = "QR",               // Type of barcode.
            BarcodeValue = "ABC123",          // Data to encode.
            ScalingFactor = "500"             // Larger factor → higher resolution.
        };

        // Generate the barcode image as a stream.
        using (Stream barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
        {
            // Ensure the stream is positioned at the beginning before insertion.
            barcodeImage.Position = 0;

            // Insert the image into the document at the current builder position.
            builder.InsertImage(barcodeImage);
        }

        // Save the document to a DOCX file.
        doc.Save("BarcodeDocument.docx");
    }
}

// Minimal stub implementation of IBarcodeGenerator.
// Replace with a real generator (e.g., one that uses Aspose.BarCode) for actual images.
class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        // For demonstration purposes return an empty stream.
        // A real implementation would generate the barcode image based on 'parameters'.
        return new MemoryStream();
    }

    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        // Return an empty stream for the legacy method as well.
        return new MemoryStream();
    }
}
