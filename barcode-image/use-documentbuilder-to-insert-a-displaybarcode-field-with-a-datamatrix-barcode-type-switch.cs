using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the custom barcode generator (required for PDF rendering and DataMatrix).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "DATAMATRIX";
        barcodeField.BarcodeValue = "ABC123";
        barcodeField.DisplayText = true; // Show the human‑readable text below the barcode.

        // Update fields to ensure the barcode is generated.
        doc.UpdateFields();

        // Save the document as PDF (the barcode will be rendered using the custom generator).
        doc.Save("output.pdf");
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator implementation.
// -----------------------------------------------------------------------------
internal class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        try
        {
            // Create the Aspose.BarCode generator with the appropriate symbology.
            BarcodeGenerator generator = new BarcodeGenerator(
                CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType),
                parameters.BarcodeValue);

            // Apply colors.
            generator.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(
                parameters.ForegroundColor, generator.Parameters.Barcode.BarColor);
            generator.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(
                parameters.BackgroundColor, generator.Parameters.BackColor);

            // Show or hide the human‑readable text.
            generator.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText
                ? CodeLocation.Below
                : CodeLocation.None;

            // Apply scaling factor if provided.
            double scalingFactor = 1.0;
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scalingFactor);

            // Adjust X‑dimension based on scaling.
            generator.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0,
                Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));

            // Apply symbol height if provided.
            if (!string.IsNullOrEmpty(parameters.SymbolHeight))
            {
                double heightInPixels = CustomBarcodeGeneratorUtils.TwipsToPixels(
                    parameters.SymbolHeight, 96, generator.Parameters.Barcode.BarHeight.Pixels);
                generator.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0, Math.Round(heightInPixels * scalingFactor));
            }

            // Apply rotation if provided.
            if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                generator.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(
                    parameters.SymbolRotation, generator.Parameters.RotationAngle);

            // Generate the barcode image into a memory stream (PNG format).
            MemoryStream ms = new MemoryStream();
            generator.Save(ms, BarCodeImageFormat.Png);
            ms.Position = 0;
            return ms;
        }
        catch
        {
            // In case of any error, return an empty stream.
            return new MemoryStream();
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // For compatibility, delegate to the primary method.
        return GetBarcodeImage(parameters);
    }
}

// -----------------------------------------------------------------------------
// Helper utilities for barcode generation.
// -----------------------------------------------------------------------------
internal static class CustomBarcodeGeneratorUtils
{
    public const double Default1DXDimensionInPixels = 1.0;

    public static SymbologyEncodeType GetBarcodeEncodeType(string barcodeType)
    {
        // Map Word barcode type strings to Aspose.BarCode symbology.
        switch (barcodeType?.ToUpperInvariant())
        {
            case "DATAMATRIX":
                return EncodeTypes.DataMatrix;
            case "QR":
                return EncodeTypes.QR;
            case "CODE128":
                return EncodeTypes.Code128;
            // Add other mappings as needed.
            default:
                return EncodeTypes.None;
        }
    }

    public static Aspose.Drawing.Color ConvertColor(string inputColor, Aspose.Drawing.Color defaultColor)
    {
        if (string.IsNullOrEmpty(inputColor))
            return defaultColor;

        try
        {
            // Input is expected as a hex string like "FF0000" (RGB).
            int argb = Convert.ToInt32(inputColor, 16);
            // Aspose.Drawing.Color expects ARGB; assume fully opaque.
            return Aspose.Drawing.Color.FromArgb(255, (argb >> 16) & 0xFF, (argb >> 8) & 0xFF, argb & 0xFF);
        }
        catch
        {
            return defaultColor;
        }
    }

    public static double ScaleFactor(string scaleFactor, double defaultValue)
    {
        if (string.IsNullOrEmpty(scaleFactor))
            return defaultValue;

        if (int.TryParse(scaleFactor, out int scale))
            return scale / 100.0;

        return defaultValue;
    }

    public static double TwipsToPixels(string heightInTwips, double resolution, double defaultValue)
    {
        if (string.IsNullOrEmpty(heightInTwips))
            return defaultValue;

        if (int.TryParse(heightInTwips, out int twips))
            return (twips / 1440.0) * resolution;

        return defaultValue;
    }

    public static float GetRotationAngle(string rotationAngle, float defaultValue)
    {
        switch (rotationAngle)
        {
            case "0": return 0f;
            case "1": return 270f;
            case "2": return 180f;
            case "3": return 90f;
            default: return defaultValue;
        }
    }
}
