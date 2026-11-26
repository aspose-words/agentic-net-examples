using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image based on the parameters supplied by the DISPLAYBARCODE field.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Map the string barcode type to a SymbologyEncodeType value.
        SymbologyEncodeType encodeType = MapBarcodeType(parameters.BarcodeType);

        // Create a memory stream to hold the PNG image.
        MemoryStream stream = new MemoryStream();

        // Generate the barcode using Aspose.BarCode.
        using (BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue))
        {
            // Apply foreground color if provided (value is a hex string like "0xFF0000").
            if (!string.IsNullOrEmpty(parameters.ForegroundColor))
                generator.Parameters.Barcode.BarColor = Color.FromArgb(Convert.ToInt32(parameters.ForegroundColor, 16));

            // Apply background color if provided.
            if (!string.IsNullOrEmpty(parameters.BackgroundColor))
                generator.Parameters.BackColor = Color.FromArgb(Convert.ToInt32(parameters.BackgroundColor, 16));

            // Save the barcode as PNG into the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Reset stream position for reading.
        stream.Position = 0;
        return stream;
    }

    // Legacy method – simply forwards to the main implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }

    // Maps the barcode type string used in the field to the corresponding EncodeTypes value.
    private SymbologyEncodeType MapBarcodeType(string barcodeType)
    {
        return barcodeType?.ToUpperInvariant() switch
        {
            "QR" => EncodeTypes.QR,
            "EAN13" => EncodeTypes.EAN13,
            "CODE128" => EncodeTypes.Code128,
            // For unsupported or unknown types, fall back to Code128.
            _ => EncodeTypes.Code128,
        };
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set static field properties.
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "PLACEHOLDER";          // Initial placeholder.
        field.ForegroundColor = "0x000000";          // Black.
        field.BackgroundColor = "0xFFFFFF";          // White.

        // Dynamic value that should replace the placeholder before barcode generation.
        string dynamicValue = "987654321";

        // Replace the placeholder with the actual value.
        field.BarcodeValue = dynamicValue;

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to trigger barcode generation.
        doc.UpdateFields();

        // Save the document as PDF (barcode will be rendered via the custom generator).
        doc.Save("output.pdf");
    }
}
