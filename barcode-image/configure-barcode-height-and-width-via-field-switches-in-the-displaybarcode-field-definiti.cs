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
        var stream = new MemoryStream();

        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Apply foreground color if supplied (value like "0xFF0000")
            if (!string.IsNullOrEmpty(parameters.ForegroundColor))
            {
                var hex = parameters.ForegroundColor.Substring(2); // remove "0x"
                generator.Parameters.Barcode.BarColor = Color.FromArgb(int.Parse(hex, System.Globalization.NumberStyles.HexNumber));
            }

            // Apply background color if supplied
            if (!string.IsNullOrEmpty(parameters.BackgroundColor))
            {
                var hex = parameters.BackgroundColor.Substring(2);
                generator.Parameters.BackColor = Color.FromArgb(int.Parse(hex, System.Globalization.NumberStyles.HexNumber));
            }

            // Save as PNG to the stream
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // For this example the old method behaves the same as the new one
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure barcode properties
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Optional visual settings
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // If you need to influence size, use the ScalingFactor switch (percentage)
        // For example, 150% makes the barcode 1.5 times larger.
        field.ScalingFactor = "150";

        // Register custom barcode generator for rendered output (PDF)
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to generate the barcode image
        doc.UpdateFields();

        // Save the document as PDF (rendered barcode)
        doc.Save("BarcodeOutput.pdf");
    }
}
