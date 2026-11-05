using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image using Aspose.BarCode and returns it as a stream.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode and background colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save the barcode as PNG into the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the main implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Input strings for which barcodes will be generated.
        string[] inputs = new string[]
        {
            "A",
            "ABCD",
            "1234567890",
            "LongerBarcodeValueExample"
        };

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        const int baseScaling = 100; // Base scaling factor (percentage).

        foreach (string text in inputs)
        {
            // Insert a new paragraph for each barcode.
            builder.Writeln();

            // Insert a DISPLAYBARCODE field using typed insertion.
            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set barcode type and value.
            field.BarcodeType = "CODE128";
            field.BarcodeValue = text;

            // Adjust scaling factor based on the length of the input string.
            int scaling = baseScaling + text.Length * 10;
            field.ScalingFactor = scaling.ToString();

            // Optional colors.
            field.ForegroundColor = "0x000000";
            field.BackgroundColor = "0xFFFFFF";
        }

        // Update fields to generate barcode images.
        doc.UpdateFields();

        // Save the document as PDF (rendered barcodes).
        doc.Save("Barcodes.pdf");
    }
}
