using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a QR code barcode field.
        InsertBarcodeField(builder, "QR", "https://example.com", displayText: true);
        builder.Writeln();

        // Insert a Code128 barcode field.
        InsertBarcodeField(builder, "CODE128", "1234567890", displayText: false);
        builder.Writeln();

        // Insert an EAN13 barcode field.
        InsertBarcodeField(builder, "EAN13", "5901234123457", displayText: true);
        builder.Writeln();

        // Generate the barcodes.
        doc.UpdateFields();

        // Save the document as PDF (rendered barcodes).
        const string outFile = "Barcodes.pdf";
        doc.Save(outFile);

        // Simple verification that the file was created.
        if (File.Exists(outFile) && new FileInfo(outFile).Length > 0)
            Console.WriteLine($"PDF saved successfully: {outFile}");
        else
            Console.WriteLine("Failed to save PDF.");
    }

    private static void InsertBarcodeField(DocumentBuilder builder, string type, string value, bool displayText)
    {
        // Insert a DISPLAYBARCODE field using the typed API.
        Field field = builder.InsertField(FieldType.FieldDisplayBarcode, false);
        if (field is FieldDisplayBarcode barcodeField)
        {
            barcodeField.BarcodeType = type;
            barcodeField.BarcodeValue = value;
            barcodeField.DisplayText = displayText;
            // Optional: set colors (hex RGB without leading #).
            barcodeField.ForegroundColor = "000000"; // black
            barcodeField.BackgroundColor = "FFFFFF"; // white
        }
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator utilities.
internal static class CustomBarcodeGeneratorUtils
{
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
        return rotationAngle switch
        {
            "0" => 0,
            "1" => 270,
            "2" => 180,
            "3" => 90,
            _ => defVal,
        };
    }

    public static QRErrorLevel GetQRCorrectionLevel(string errorCorrectionLevel, QRErrorLevel def)
    {
        return errorCorrectionLevel switch
        {
            "0" => QRErrorLevel.LevelL,
            "1" => QRErrorLevel.LevelM,
            "2" => QRErrorLevel.LevelQ,
            "3" => QRErrorLevel.LevelH,
            _ => def,
        };
    }

    public static SymbologyEncodeType GetBarcodeEncodeType(string encodeTypeFromWord)
    {
        return encodeTypeFromWord switch
        {
            "QR" => EncodeTypes.QR,
            "CODE128" => EncodeTypes.Code128,
            "EAN13" => EncodeTypes.EAN13,
            "EAN8" => EncodeTypes.EAN8,
            "UPCA" => EncodeTypes.UPCA,
            "UPCE" => EncodeTypes.UPCE,
            "NW7" => EncodeTypes.Codabar,
            _ => EncodeTypes.None,
        };
    }

    public static Aspose.Drawing.Color ConvertColor(string inputColor, Aspose.Drawing.Color defVal)
    {
        if (string.IsNullOrEmpty(inputColor))
            return defVal;
        try
        {
            int color = Convert.ToInt32(inputColor, 16);
            // Input is assumed to be RGB hex (e.g., "FF0000" for red).
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

    public const double DefaultQRXDimensionInPixels = 4.0;
    public const double Default1DXDimensionInPixels = 1.0;

    public static Aspose.Drawing.Bitmap DrawErrorImage(Exception error)
    {
        var bmp = new Aspose.Drawing.Bitmap(200, 50);
        using (var grf = Aspose.Drawing.Graphics.FromImage(bmp))
        {
            grf.Clear(Aspose.Drawing.Color.White);
            using var font = new Aspose.Drawing.Font("Arial", 8f, Aspose.Drawing.FontStyle.Regular);
            grf.DrawString(error.Message, font, Aspose.Drawing.Brushes.Red, new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
        }
        return bmp;
    }

    public static Stream ConvertImageToWord(Aspose.Drawing.Bitmap bmp)
    {
        var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        ms.Position = 0;
        return ms;
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator implementation.
internal class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        try
        {
            var generator = new BarcodeGenerator(
                CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType),
                parameters.BarcodeValue);

            // Colors.
            generator.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.ForegroundColor, generator.Parameters.Barcode.BarColor);
            generator.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.BackgroundColor, generator.Parameters.BackColor);

            // Display text.
            generator.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText
                ? CodeLocation.Below
                : CodeLocation.None;

            // QR error correction.
            generator.Parameters.Barcode.QR.ErrorLevel = QRErrorLevel.LevelH;
            if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel))
                generator.Parameters.Barcode.QR.ErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(parameters.ErrorCorrectionLevel, generator.Parameters.Barcode.QR.ErrorLevel);

            // Rotation.
            if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                generator.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(parameters.SymbolRotation, generator.Parameters.RotationAngle);

            // Scaling.
            double scalingFactor = 1;
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scalingFactor);

            // X dimension.
            if (generator.BarcodeType == EncodeTypes.QR)
                generator.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scalingFactor));
            else
                generator.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));

            // Height.
            if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                generator.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0,
                    Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(parameters.SymbolHeight, 96, generator.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

            // PosCodeStyle handling can be added here if needed.

            // Generate image.
            using var img = generator.GenerateBarCodeImage();
            var ms = new MemoryStream();
            img.Save(ms, ImageFormat.Png);
            ms.Position = 0;
            return ms;
        }
        catch (Exception ex)
        {
            // Return an error image if generation fails.
            var errBmp = CustomBarcodeGeneratorUtils.DrawErrorImage(ex);
            return CustomBarcodeGeneratorUtils.ConvertImageToWord(errBmp);
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}
