using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the cached barcode generator.
        doc.FieldOptions.BarcodeGenerator = new CachedBarcodeGenerator();

        // First barcode – parameters will be cached after generation.
        FieldDisplayBarcode field1 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field1.BarcodeType = "QR";
        field1.BarcodeValue = "CACHE123";
        field1.BackgroundColor = "0xF8BD69";
        field1.ForegroundColor = "0xB5413B";
        field1.ErrorCorrectionLevel = "3";
        field1.ScalingFactor = "250";
        field1.SymbolHeight = "1000";
        field1.SymbolRotation = "0";

        builder.Writeln();

        // Second barcode – identical parameters, should be served from cache.
        FieldDisplayBarcode field2 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field2.BarcodeType = "QR";
        field2.BarcodeValue = "CACHE123";
        field2.BackgroundColor = "0xF8BD69";
        field2.ForegroundColor = "0xB5413B";
        field2.ErrorCorrectionLevel = "3";
        field2.ScalingFactor = "250";
        field2.SymbolHeight = "1000";
        field2.SymbolRotation = "0";

        // Update fields to trigger barcode generation.
        doc.UpdateFields();

        // Save as PDF – a rendered format that forces the custom generator to run.
        doc.Save("Barcodes.pdf");
    }
}

// -----------------------------------------------------------------------------
// Cached barcode generator – stores generated images for identical parameter sets.
// -----------------------------------------------------------------------------
internal class CachedBarcodeGenerator : IBarcodeGenerator
{
    private readonly CustomBarcodeGenerator _inner = new CustomBarcodeGenerator();
    private static readonly Dictionary<string, byte[]> _cache = new Dictionary<string, byte[]>();

    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        string key = GenerateKey(parameters);
        if (_cache.TryGetValue(key, out var cachedData))
        {
            // Return a fresh stream based on the cached byte array.
            return new MemoryStream(cachedData) { Position = 0 };
        }

        // Generate the image, cache it, and return a new stream.
        using (Stream generated = _inner.GetBarcodeImage(parameters))
        using (MemoryStream ms = new MemoryStream())
        {
            generated.CopyTo(ms);
            byte[] data = ms.ToArray();
            _cache[key] = data;
            return new MemoryStream(data) { Position = 0 };
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // For legacy fields we delegate to the same implementation.
        return GetBarcodeImage(parameters);
    }

    private static string GenerateKey(Aspose.Words.Fields.BarcodeParameters p)
    {
        // Concatenate all properties that influence the visual output.
        return string.Join("|",
            p.BarcodeType ?? "",
            p.BarcodeValue ?? "",
            p.BackgroundColor ?? "",
            p.ForegroundColor ?? "",
            p.ErrorCorrectionLevel ?? "",
            p.ScalingFactor ?? "",
            p.SymbolHeight ?? "",
            p.SymbolRotation ?? "",
            p.DisplayText.ToString(),
            p.PosCodeStyle ?? "",
            p.CaseCodeStyle ?? "",
            p.FixCheckDigit.ToString());
    }
}

// -----------------------------------------------------------------------------
// Helper utilities used by the custom barcode generator.
// -----------------------------------------------------------------------------
internal static class CustomBarcodeGeneratorUtils
{
    public static double TwipsToPixels(string heightInTwips, double defVal)
    {
        return TwipsToPixels(heightInTwips, 96, defVal);
    }

    public static double TwipsToPixels(string heightInTwips, double resolution, double defVal)
    {
        try
        {
            int lVal = int.Parse(heightInTwips);
            return (lVal / 1440.0) * resolution;
        }
        catch
        {
            return defVal;
        }
    }

    public static float GetRotationAngle(string rotationAngle, float defVal)
    {
        switch (rotationAngle)
        {
            case "0": return 0;
            case "1": return 270;
            case "2": return 180;
            case "3": return 90;
            default: return defVal;
        }
    }

    public static QRErrorLevel GetQRCorrectionLevel(string errorCorrectionLevel, QRErrorLevel def)
    {
        switch (errorCorrectionLevel)
        {
            case "0": return QRErrorLevel.LevelL;
            case "1": return QRErrorLevel.LevelM;
            case "2": return QRErrorLevel.LevelQ;
            case "3": return QRErrorLevel.LevelH;
            default: return def;
        }
    }

    public static SymbologyEncodeType GetBarcodeEncodeType(string encodeTypeFromWord)
    {
        switch (encodeTypeFromWord)
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

    public static Aspose.Drawing.Color ConvertColor(string inputColor, Aspose.Drawing.Color defVal)
    {
        if (string.IsNullOrEmpty(inputColor)) return defVal;
        try
        {
            int color = Convert.ToInt32(inputColor, 16);
            return Aspose.Drawing.Color.FromArgb(color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF);
        }
        catch
        {
            return defVal;
        }
    }

    public static double ScaleFactor(string scaleFactor, double defVal)
    {
        try
        {
            int scale = int.Parse(scaleFactor);
            return scale / 100.0;
        }
        catch
        {
            return defVal;
        }
    }

    public static void SetPosCodeStyle(BarcodeGenerator gen, string posCodeStyle, string barcodeValue)
    {
        switch (posCodeStyle)
        {
            case "SUP2":
                gen.CodeText = barcodeValue.Substring(0, barcodeValue.Length - 2);
                gen.Parameters.Barcode.Supplement.SupplementData = barcodeValue.Substring(barcodeValue.Length - 2, 2);
                break;
            case "SUP5":
                gen.CodeText = barcodeValue.Substring(0, barcodeValue.Length - 5);
                gen.Parameters.Barcode.Supplement.SupplementData = barcodeValue.Substring(barcodeValue.Length - 5, 5);
                break;
            case "CASE":
                gen.Parameters.Border.Visible = true;
                gen.Parameters.Border.Color = gen.Parameters.Barcode.BarColor;
                gen.Parameters.Border.DashStyle = BorderDashStyle.Solid;
                gen.Parameters.Border.Width.Pixels = gen.Parameters.Barcode.XDimension.Pixels * 5;
                break;
        }
    }

    public const double DefaultQRXDimensionInPixels = 4.0;
    public const double Default1DXDimensionInPixels = 1.0;

    public static Aspose.Drawing.Bitmap DrawErrorImage(Exception error)
    {
        Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(100, 100);
        using (Aspose.Drawing.Graphics grf = Aspose.Drawing.Graphics.FromImage(bmp))
        {
            Aspose.Drawing.Font font = new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, Aspose.Drawing.FontStyle.Regular);
            grf.DrawString(error.Message, font, Aspose.Drawing.Brushes.Red, new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
        }
        return bmp;
    }

    public static Stream ConvertImageToWord(Aspose.Drawing.Bitmap bmp)
    {
        MemoryStream ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        ms.Position = 0;
        return ms;
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator that creates barcode images using Aspose.BarCode.
// -----------------------------------------------------------------------------
internal class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        try
        {
            BarcodeGenerator gen = new BarcodeGenerator(
                CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType),
                parameters.BarcodeValue);

            gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(
                parameters.ForegroundColor, gen.Parameters.Barcode.BarColor);
            gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(
                parameters.BackgroundColor, gen.Parameters.BackColor);

            gen.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText
                ? CodeLocation.Below
                : CodeLocation.None;

            gen.Parameters.Barcode.QR.ErrorLevel = QRErrorLevel.LevelH;
            if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel))
                gen.Parameters.Barcode.QR.ErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(
                    parameters.ErrorCorrectionLevel, gen.Parameters.Barcode.QR.ErrorLevel);

            if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                gen.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(
                    parameters.SymbolRotation, gen.Parameters.RotationAngle);

            double scalingFactor = 1;
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(
                    parameters.ScalingFactor, scalingFactor);

            if (gen.BarcodeType == EncodeTypes.QR)
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0,
                    Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scalingFactor));
            else
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0,
                    Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));

            if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0,
                    Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(
                        parameters.SymbolHeight, gen.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

            if (!string.IsNullOrEmpty(parameters.PosCodeStyle))
                CustomBarcodeGeneratorUtils.SetPosCodeStyle(gen, parameters.PosCodeStyle, parameters.BarcodeValue);

            return CustomBarcodeGeneratorUtils.ConvertImageToWord(gen.GenerateBarCodeImage());
        }
        catch (Exception e)
        {
            return CustomBarcodeGeneratorUtils.ConvertImageToWord(
                CustomBarcodeGeneratorUtils.DrawErrorImage(e));
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Delegate to the main method.
        return GetBarcodeImage(parameters);
    }
}
