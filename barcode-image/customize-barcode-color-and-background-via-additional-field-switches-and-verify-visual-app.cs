using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Create a barcode generator using the requested symbology.
        var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue);

        // Apply foreground (barcode) color if provided.
        if (!string.IsNullOrEmpty(parameters.ForegroundColor))
            generator.Parameters.Barcode.BarColor = ParseColor(parameters.ForegroundColor);

        // Apply background color if provided.
        if (!string.IsNullOrEmpty(parameters.BackgroundColor))
            generator.Parameters.BackColor = ParseColor(parameters.BackgroundColor);

        // Save the barcode image to a memory stream in PNG format.
        var stream = new MemoryStream();
        generator.Save(stream, BarCodeImageFormat.Png);
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the current implementation.
        return GetBarcodeImage(parameters);
    }

    private static Color ParseColor(string colorString)
    {
        // Expected format: "0xRRGGBB" (hex). Remove the "0x" prefix if present.
        if (colorString.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            colorString = colorString.Substring(2);

        // Parse the hex value.
        int argb = int.Parse(colorString, System.Globalization.NumberStyles.HexNumber);
        // The value is RGB; convert to ARGB with full opacity.
        return Color.FromArgb(255, (argb >> 16) & 0xFF, (argb >> 8) & 0xFF, argb & 0xFF);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Customize colors via field switches.
        field.ForegroundColor = "0xFF0000"; // Red barcode bars.
        field.BackgroundColor = "0xFFFF00"; // Yellow background.

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to generate the barcode image.
        doc.UpdateFields();

        // Save the document as PDF.
        const string outputPath = "barcode.pdf";
        doc.Save(outputPath);

        // Simple verification: ensure the PDF file was created and has content.
        if (File.Exists(outputPath) && new FileInfo(outputPath).Length > 0)
        {
            Console.WriteLine($"PDF generated successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to generate PDF.");
        }
    }
}
