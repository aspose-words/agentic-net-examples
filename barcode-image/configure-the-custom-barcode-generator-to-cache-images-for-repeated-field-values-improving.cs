using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Simple in‑memory cache to avoid regenerating the same barcode image.
    private static readonly Dictionary<string, byte[]> _cache = new Dictionary<string, byte[]>();

    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Build a cache key from type and value.
        string key = $"{parameters.BarcodeType}|{parameters.BarcodeValue}";
        if (_cache.TryGetValue(key, out var cachedBytes))
        {
            // Return a fresh stream based on the cached bytes.
            return new MemoryStream(cachedBytes);
        }

        // Generate the barcode image.
        using (BarcodeGenerator generator = new BarcodeGenerator(MapEncodeType(parameters.BarcodeType), parameters.BarcodeValue))
        {
            // Example of setting colors (optional).
            // generator.Parameters.Barcode.BarColor = Color.Black;
            // generator.Parameters.BackColor = Color.White;

            using (MemoryStream ms = new MemoryStream())
            {
                generator.Save(ms, BarCodeImageFormat.Png);
                byte[] bytes = ms.ToArray();

                // Store in cache for future calls.
                _cache[key] = bytes;

                // Return a new stream positioned at the beginning.
                return new MemoryStream(bytes);
            }
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Legacy method simply forwards to the current implementation.
        return GetBarcodeImage(parameters);
    }

    // Maps the string barcode type used in Word fields to the Aspose.BarCode SymbologyEncodeType.
    private static SymbologyEncodeType MapEncodeType(string barcodeType)
    {
        switch (barcodeType?.ToUpperInvariant())
        {
            case "CODE128":
                return EncodeTypes.Code128;
            case "QR":
                return EncodeTypes.QR;
            case "CODE39":
                return EncodeTypes.Code39;
            case "EAN13":
                return EncodeTypes.EAN13;
            default:
                // Fallback to a supported type.
                return EncodeTypes.Code128;
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first barcode field.
        builder.Writeln("First barcode:");
        FieldDisplayBarcode field1 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field1.BarcodeType = "CODE128";
        field1.BarcodeValue = "1234567890";

        // Insert a second barcode field with the same value to demonstrate caching.
        builder.Writeln();
        builder.Writeln("Second barcode (same value):");
        FieldDisplayBarcode field2 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field2.BarcodeType = "CODE128";
        field2.BarcodeValue = "1234567890";

        // Register the custom barcode generator so that rendered outputs (PDF) use it.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render the barcode fields.
        doc.UpdateFields();

        // Save the document as PDF, which will contain the rendered barcode images.
        doc.Save("output.pdf");
    }
}
