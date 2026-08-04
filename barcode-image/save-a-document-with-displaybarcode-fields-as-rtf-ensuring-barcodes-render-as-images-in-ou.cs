using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.BarCode.Generation;

public class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image based on the supplied parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Convert the barcode type string to the corresponding SymbologyEncodeType enum value.
        SymbologyEncodeType encodeType = (SymbologyEncodeType)Enum.Parse(
            typeof(SymbologyEncodeType), parameters.BarcodeType, true);

        // Create the barcode generator.
        var generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue);

        // Optional: apply background color if provided.
        if (!string.IsNullOrEmpty(parameters.BackgroundColor))
        {
            // Expected format: "0xRRGGBB"
            int argb = int.Parse(parameters.BackgroundColor.Substring(2),
                System.Globalization.NumberStyles.HexNumber);
            generator.Parameters.BackColor = Aspose.Drawing.Color.FromArgb(argb);
        }

        // Optional: apply foreground color if provided.
        if (!string.IsNullOrEmpty(parameters.ForegroundColor))
        {
            int argb = int.Parse(parameters.ForegroundColor.Substring(2),
                System.Globalization.NumberStyles.HexNumber);
            // The property name may be ForeColor in newer versions; use reflection for compatibility.
            var foreColorProp = generator.Parameters.GetType().GetProperty("ForeColor");
            if (foreColorProp != null)
                foreColorProp.SetValue(generator.Parameters, Aspose.Drawing.Color.FromArgb(argb));
        }

        // Generate the image into a memory stream.
        var ms = new MemoryStream();
        generator.Save(ms, BarCodeImageFormat.Png);
        ms.Position = 0;
        return ms;
    }

    // Legacy method simply forwards to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        var doc = new Document();

        // Register the custom barcode generator so that DISPLAYBARCODE fields are rendered as images.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a DISPLAYBARCODE field using the typed API.
        var builder = new DocumentBuilder(doc);
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set barcode properties.
        field.BarcodeType = "QR";
        field.BarcodeValue = "ABC123";
        field.BackgroundColor = "0xF8BD69";
        field.ForegroundColor = "0xB5413B";
        field.ScalingFactor = "250";

        // Update fields to generate the barcode image.
        doc.UpdateFields();

        // Save the document as RTF; the barcode will be stored as an image.
        var rtfOptions = new RtfSaveOptions();
        doc.Save("Barcodes.rtf", rtfOptions);
    }
}
