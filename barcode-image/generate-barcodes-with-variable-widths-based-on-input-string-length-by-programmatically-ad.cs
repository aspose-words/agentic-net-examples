using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image using Aspose.BarCode and returns it as a stream.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        MemoryStream stream = new MemoryStream();
        // Use Code128 symbology; the barcode value comes from the field parameters.
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Optional: set colors if needed (commented out because BarcodeParameters does not expose color properties).
            // generator.Parameters.Barcode.BarColor = Color.Black;
            // generator.Parameters.BackColor = Color.White;

            generator.Save(stream, BarCodeImageFormat.Png);
        }
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
        // Input strings for which barcodes will be generated.
        string[] values = new[] { "A", "AB", "ABC", "ABCD", "ABCDE" };

        // Determine the maximum length to calculate scaling factors.
        int maxLength = 0;
        foreach (var v in values)
            if (v.Length > maxLength) maxLength = v.Length;

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Generate a DISPLAYBARCODE field for each input string.
        foreach (var value in values)
        {
            // Insert a new paragraph for each barcode.
            builder.Writeln();

            // Insert a typed DISPLAYBARCODE field.
            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "CODE128";
            field.BarcodeValue = value;

            // Compute scaling factor based on length (range 80–150).
            int scale = 80 + (value.Length * 70) / maxLength; // simple linear scaling.
            field.ScalingFactor = scale.ToString();

            // Set colors (black on white) using the field's string color properties.
            field.ForegroundColor = "0x000000";
            field.BackgroundColor = "0xFFFFFF";
        }

        // Render the fields so that the custom generator creates the barcode images.
        doc.UpdateFields();

        // Save the document as PDF (rendered barcodes).
        doc.Save("Barcodes.pdf");
    }
}
