using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;
using Aspose.BarCode;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Simple RTF content
        string rtfContent = @"{\rtf1\ansi This is a sample RTF document.\par}";
        using (MemoryStream rtfStream = new MemoryStream(Encoding.UTF8.GetBytes(rtfContent)))
        {
            // Load the RTF document
            Document doc = new Document(rtfStream);

            // Register the custom barcode generator
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Insert a barcode field
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeType = "CODE128";
            barcodeField.BarcodeValue = "1234567890";
            barcodeField.DisplayText = true;

            // Foreground and background colors must be supplied as hex strings (RRGGBB)
            barcodeField.ForegroundColor = $"{Color.Black.R:X2}{Color.Black.G:X2}{Color.Black.B:X2}";
            barcodeField.BackgroundColor = $"{Color.White.R:X2}{Color.White.G:X2}{Color.White.B:X2}";

            barcodeField.ScalingFactor = "100"; // 100% scaling

            // Update fields to apply the barcode generator
            doc.UpdateFields();

            // Save as DOCX
            doc.Save("Output.docx");
        }
    }
}

internal static class CustomBarcodeGeneratorUtils
{
    public static SymbologyEncodeType GetBarcodeEncodeType(string encodeTypeFromWord)
    {
        switch (encodeTypeFromWord?.ToUpperInvariant())
        {
            case "QR": return EncodeTypes.QR;
            case "CODE128": return EncodeTypes.Code128;
            case "JPPOST": return EncodeTypes.RM4SCC;
            case "EAN8":
            case "JAN8": return EncodeTypes.EAN8;
            case "EAN13":
            case "JAN13": return EncodeTypes.EAN13;
            case "UPCA": return EncodeTypes.UPCA;
            case "UPCE": return EncodeTypes.UPCE;
            case "CASE":
            case "ITF14": return EncodeTypes.ITF14;
            case "NW7": return EncodeTypes.Codabar;
            default: return EncodeTypes.None;
        }
    }

    public static Color ConvertColor(string inputColor, Color defVal)
    {
        if (string.IsNullOrEmpty(inputColor))
            return defVal;

        try
        {
            // Expecting hex string like "FF0000" (RGB)
            int rgb = Convert.ToInt32(inputColor, 16);
            return Color.FromArgb(255, (rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
        }
        catch
        {
            return defVal;
        }
    }

    public static double ScaleFactor(string scaleFactor, double defVal)
    {
        if (string.IsNullOrEmpty(scaleFactor))
            return defVal;

        if (int.TryParse(scaleFactor, out int scale))
            return scale / 100.0;

        return defVal;
    }

    public const double DefaultQRXDimensionInPixels = 4.0;
    public const double Default1DXDimensionInPixels = 1.0;
}

internal class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        try
        {
            // Determine the barcode type
            SymbologyEncodeType encodeType = CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType);
            BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue);

            // Colors
            generator.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.ForegroundColor, generator.Parameters.Barcode.BarColor);
            generator.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.BackgroundColor, generator.Parameters.BackColor);

            // Display text
            generator.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText
                ? CodeLocation.Below
                : CodeLocation.None;

            // Scaling
            double scaling = 1.0;
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                scaling = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scaling);

            // X dimension (module size)
            if (generator.BarcodeType == EncodeTypes.QR)
                generator.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scaling));
            else
                generator.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scaling));

            // Generate image
            using (Image img = generator.GenerateBarCodeImage())
            {
                MemoryStream ms = new MemoryStream();
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                return ms;
            }
        }
        catch
        {
            // Return an empty stream on error
            return new MemoryStream();
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the current implementation
        return GetBarcodeImage(parameters);
    }
}
