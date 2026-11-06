using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Simple in‑memory logger.
    private static readonly List<string> _log = new List<string>();

    public static IReadOnlyList<string> Log => _log.AsReadOnly();

    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Generate barcode image using Aspose.BarCode.
        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Optional: set colors if needed.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Record image size (bytes) and barcode value (used as field identifier).
        long imageSize = stream.Length;
        string entry = $"BarcodeValue: {parameters.BarcodeValue}, ImageSizeBytes: {imageSize}";
        _log.Add(entry);

        // Reset stream position for the caller.
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the current implementation.
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First barcode field.
        FieldDisplayBarcode field1 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field1.BarcodeType = "CODE128";
        field1.BarcodeValue = "1234567890";
        field1.ForegroundColor = "0x000000";
        field1.BackgroundColor = "0xFFFFFF";

        // Insert a paragraph break.
        builder.Writeln();

        // Second barcode field.
        FieldDisplayBarcode field2 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field2.BarcodeType = "QR";
        field2.BarcodeValue = "HelloWorld";
        field2.ForegroundColor = "0x000000";
        field2.BackgroundColor = "0xFFFFFF";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to trigger barcode generation.
        doc.UpdateFields();

        // Save the document as PDF (rendered barcodes).
        doc.Save("BarcodesOutput.pdf");

        // Output the logging information.
        Console.WriteLine("Barcode generation log:");
        foreach (string entry in CustomBarcodeGenerator.Log)
        {
            Console.WriteLine(entry);
        }
    }
}
