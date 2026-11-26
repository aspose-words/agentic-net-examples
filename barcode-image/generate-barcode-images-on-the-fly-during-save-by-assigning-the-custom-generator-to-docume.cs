using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Map the string barcode type to an EncodeTypes value.
        SymbologyEncodeType encodeType = GetEncodeType(parameters.BarcodeType);

        // Create the barcode generator with the mapped type and the barcode text.
        using var generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue);

        // Optional: set barcode colors if the parameters contain color strings.
        // The color strings are expected in hexadecimal format, e.g., "0xFF0000".
        if (!string.IsNullOrEmpty(parameters.ForegroundColor))
        {
            int argb = Convert.ToInt32(parameters.ForegroundColor.Replace("0x", ""), 16);
            generator.Parameters.Barcode.BarColor = Color.FromArgb(argb);
        }

        if (!string.IsNullOrEmpty(parameters.BackgroundColor))
        {
            int argb = Convert.ToInt32(parameters.BackgroundColor.Replace("0x", ""), 16);
            generator.Parameters.BackColor = Color.FromArgb(argb);
        }

        // Save the generated barcode image to a memory stream in PNG format.
        var stream = new MemoryStream();
        generator.Save(stream, BarCodeImageFormat.Png);
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the primary implementation.
        return GetBarcodeImage(parameters);
    }

    private SymbologyEncodeType GetEncodeType(string type)
    {
        // Default to Code128 if the type is unknown.
        return type?.ToUpperInvariant() switch
        {
            "CODE128" => EncodeTypes.Code128,
            "QR" => EncodeTypes.QR,
            "CODE39" => EncodeTypes.Code39,
            "EAN13" => EncodeTypes.EAN13,
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
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Register the custom barcode generator for rendered outputs (PDF in this case).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to trigger barcode generation.
        doc.UpdateFields();

        // Save the document as PDF; the barcode image will be rendered via the custom generator.
        doc.Save("output.pdf");
    }
}
