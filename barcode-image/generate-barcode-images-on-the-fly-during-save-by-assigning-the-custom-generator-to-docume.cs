using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image based on the provided parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();

        // Use Aspose.BarCode to generate the barcode.
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save the barcode as PNG to the memory stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Reset stream position for reading.
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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";               // Set barcode symbology.
        field.BarcodeValue = "1234567890";           // Set barcode text.
        field.ForegroundColor = "0x000000";         // Black foreground.
        field.BackgroundColor = "0xFFFFFF";         // White background.

        // Register the custom barcode generator for rendered outputs (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document as PDF, which will invoke the custom generator.
        doc.Save("output.pdf");
    }
}
