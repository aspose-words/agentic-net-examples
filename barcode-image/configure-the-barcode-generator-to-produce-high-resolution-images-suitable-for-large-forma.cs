using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a high‑resolution PNG barcode image based on the supplied parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Map the string barcode type from Word to Aspose.BarCode EncodeTypes.
        SymbologyEncodeType encodeType = GetEncodeType(parameters.BarcodeType);

        // Create the barcode generator with the desired symbology and value.
        using (BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue))
        {
            // Set colors explicitly.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save the barcode to a memory stream as PNG.
            MemoryStream stream = new MemoryStream();
            generator.Save(stream, BarCodeImageFormat.Png);
            stream.Position = 0;
            return stream;
        }
    }

    // Legacy method delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }

    // Helper to convert Word barcode type strings to EncodeTypes.
    private static SymbologyEncodeType GetEncodeType(string type)
    {
        return type?.ToUpperInvariant() switch
        {
            "CODE128" => EncodeTypes.Code128,
            "QR" => EncodeTypes.QR,
            _ => EncodeTypes.Code128 // Default fallback.
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

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields so the barcode is rendered.
        doc.UpdateFields();

        // Save the document as a high‑resolution PDF.
        doc.Save("output.pdf");
    }
}
