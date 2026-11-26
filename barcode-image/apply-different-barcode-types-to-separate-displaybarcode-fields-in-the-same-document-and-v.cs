using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Stores generated barcode values for later verification.
    public static List<string> GeneratedBarcodes { get; } = new List<string>();

    // Generates a PNG image for the given barcode parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Map the string BarcodeType to a SymbologyEncodeType using EncodeTypes static values.
        SymbologyEncodeType encodeType = parameters.BarcodeType switch
        {
            "QR" => EncodeTypes.QR,
            "CODE128" => EncodeTypes.Code128,
            "CODE39" => EncodeTypes.Code39,
            "EAN13" => EncodeTypes.EAN13,
            _ => EncodeTypes.Code128
        };

        // Record the barcode value for verification.
        GeneratedBarcodes.Add(parameters.BarcodeValue);

        // Generate the barcode image.
        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue))
        {
            // Set barcode and background colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save as PNG to the stream.
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

        // First barcode: QR code.
        FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "https://example.com";
        qrField.ForegroundColor = "0x000000";
        qrField.BackgroundColor = "0xFFFFFF";

        // Paragraph break between barcodes.
        builder.Writeln();

        // Second barcode: CODE128.
        FieldDisplayBarcode code128Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        code128Field.BarcodeType = "CODE128";
        code128Field.BarcodeValue = "1234567890";
        code128Field.ForegroundColor = "0x0000FF";
        code128Field.BackgroundColor = "0xFFFFFF";

        builder.Writeln();

        // Third barcode: CODE39.
        FieldDisplayBarcode code39Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "CODE39TEST";
        code39Field.ForegroundColor = "0xFF0000";
        code39Field.BackgroundColor = "0xFFFFFF";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to generate barcode images.
        doc.UpdateFields();

        // Save the document as PDF (rendered barcodes).
        doc.Save("output.pdf");

        // Verify that all barcodes were generated.
        int expectedCount = 3;
        if (CustomBarcodeGenerator.GeneratedBarcodes.Count == expectedCount)
        {
            Console.WriteLine("All barcodes rendered successfully.");
        }
        else
        {
            Console.WriteLine($"Expected {expectedCount} barcodes, but generated {CustomBarcodeGenerator.GeneratedBarcodes.Count}.");
        }
    }
}
