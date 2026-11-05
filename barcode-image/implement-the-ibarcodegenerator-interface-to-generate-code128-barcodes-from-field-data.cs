using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image using Aspose.BarCode and returns it as a PNG stream.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        // EncodeTypes.Code128 is the correct static value for Code128 symbology.
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Save the generated barcode as PNG into the memory stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        // Reset the stream position so it can be read from the beginning.
        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a typed DISPLAYBARCODE field.
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";               // Set barcode symbology.
        field.BarcodeValue = "1234567890";           // The data to encode.
        field.ForegroundColor = "0x000000";         // Black foreground.
        field.BackgroundColor = "0xFFFFFF";         // White background.

        // Register the custom barcode generator for rendered outputs (PDF in this case).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document as PDF; the barcode will be rendered using the custom generator.
        doc.Save("output.pdf");
    }
}
