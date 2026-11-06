using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image based on the provided parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        MemoryStream stream = new MemoryStream();

        // Map the string barcode type to a SymbologyEncodeType value.
        SymbologyEncodeType encodeType = MapBarcodeType(parameters.BarcodeType);

        using (BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue))
        {
            // Apply foreground color if supplied.
            if (!string.IsNullOrEmpty(parameters.ForegroundColor))
            {
                // Remove possible "0x" prefix before conversion.
                string hex = parameters.ForegroundColor.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                    ? parameters.ForegroundColor.Substring(2)
                    : parameters.ForegroundColor;
                int argb = Convert.ToInt32(hex, 16);
                generator.Parameters.Barcode.BarColor = Color.FromArgb(argb);
            }

            // Apply background color if supplied.
            if (!string.IsNullOrEmpty(parameters.BackgroundColor))
            {
                string hex = parameters.BackgroundColor.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                    ? parameters.BackgroundColor.Substring(2)
                    : parameters.BackgroundColor;
                int argb = Convert.ToInt32(hex, 16);
                generator.Parameters.BackColor = Color.FromArgb(argb);
            }

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

    // Maps the string representation of a barcode type to the corresponding SymbologyEncodeType.
    private static SymbologyEncodeType MapBarcodeType(string type)
    {
        if (string.IsNullOrEmpty(type))
            return EncodeTypes.Code128;

        switch (type.Trim().ToUpperInvariant())
        {
            case "QR":
                return EncodeTypes.QR;
            case "CODE128":
                return EncodeTypes.Code128;
            case "CODE39":
                return EncodeTypes.Code39;
            case "EAN13":
                return EncodeTypes.EAN13;
            default:
                return EncodeTypes.Code128;
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First barcode field – QR code.
        FieldDisplayBarcode fieldQr = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        fieldQr.BarcodeType = "QR";
        fieldQr.BarcodeValue = "https://example.com";
        fieldQr.ForegroundColor = "0x000000";
        fieldQr.BackgroundColor = "0xFFFFFF";

        // Second barcode field – CODE128.
        FieldDisplayBarcode fieldCode128 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        fieldCode128.BarcodeType = "CODE128";
        fieldCode128.BarcodeValue = "1234567890";
        fieldCode128.ForegroundColor = "0x0000FF";
        fieldCode128.BackgroundColor = "0xFFFFFF";

        // Third barcode field – CODE39.
        FieldDisplayBarcode fieldCode39 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        fieldCode39.BarcodeType = "CODE39";
        fieldCode39.BarcodeValue = "CODE39TEST";
        fieldCode39.ForegroundColor = "0xFF0000";
        fieldCode39.BackgroundColor = "0xFFFFFF";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Render the barcode fields.
        doc.UpdateFields();

        // Save the document as PDF, which will contain rendered barcodes.
        string outputPath = "BarcodesOutput.pdf";
        doc.Save(outputPath);

        // Verification: generate barcode images directly via the custom generator.
        CustomBarcodeGenerator verifier = new CustomBarcodeGenerator();

        var parametersList = new[]
        {
            new Aspose.Words.Fields.BarcodeParameters { BarcodeType = "QR", BarcodeValue = "https://example.com" },
            new Aspose.Words.Fields.BarcodeParameters { BarcodeType = "CODE128", BarcodeValue = "1234567890" },
            new Aspose.Words.Fields.BarcodeParameters { BarcodeType = "CODE39", BarcodeValue = "CODE39TEST" }
        };

        foreach (var parms in parametersList)
        {
            using (Stream imgStream = verifier.GetBarcodeImage(parms))
            {
                Console.WriteLine($"Generated barcode type {parms.BarcodeType} image size: {imgStream.Length} bytes");
            }
        }

        Console.WriteLine($"Document saved to {outputPath}");
    }
}
