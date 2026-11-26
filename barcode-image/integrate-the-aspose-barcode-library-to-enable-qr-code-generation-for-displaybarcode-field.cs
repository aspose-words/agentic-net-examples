using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a QR code image using Aspose.BarCode and returns it as a PNG stream.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();

        // Use QR encoding type and the value supplied by the field.
        using (var generator = new BarcodeGenerator(EncodeTypes.QR, parameters.BarcodeValue))
        {
            // Set barcode colors.
            generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            generator.Parameters.BackColor = Aspose.Drawing.Color.White;

            // Save the barcode as PNG into the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Reset the stream position for reading.
        stream.Position = 0;
        return stream;
    }

    // Legacy method – delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field and configure it as a QR code.
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "QR";
        field.BarcodeValue = "https://example.com";
        field.ForegroundColor = "0x000000"; // Black
        field.BackgroundColor = "0xFFFFFF"; // White

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render the barcode.
        doc.UpdateFields();

        // Save the document as PDF with the rendered QR code.
        doc.Save("QrBarcodeOutput.pdf");
    }
}
