using System;
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
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the custom barcode generator – required for PDF rendering.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "ABC123";
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Generate the barcode image.
        doc.UpdateFields();

        // Save the document as PDF.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomBarcode.pdf");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Optional: output the field code for verification.
        Console.WriteLine("Field code:");
        Console.WriteLine(barcodeField.GetFieldCode());
    }
}

// -----------------------------------------------------------------------------
// Helper utilities for the custom barcode generator.
// -----------------------------------------------------------------------------
internal class CustomBarcodeGeneratorUtils
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
            "JPPOST" => EncodeTypes.RM4SCC,
            "EAN8" => EncodeTypes.EAN8,
            "JAN8" => EncodeTypes.EAN8,
            "EAN13" => EncodeTypes.EAN13,
            "JAN13" => EncodeTypes.EAN13,
            "UPCA" => EncodeTypes.UPCA,
            "UPCE" => EncodeTypes.UPCE,
            "CASE" => EncodeTypes.ITF14,
            "ITF14" => EncodeTypes.ITF14,
            "NW7" => EncodeTypes.Codabar,
            _ => EncodeTypes.None,
        };
    }

    public static Color ConvertColor(string inputColor, Color defVal)
    {
        if (string.IsNullOrEmpty(inputColor))
            return defVal;
        try
        {
            int color = Convert.ToInt32(inputColor, 16);
            return Color.FromArgb(color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF);
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

    public static Bitmap DrawErrorImage(Exception error)
    {
        Bitmap bmp = new Bitmap(100, 100);
        using (Graphics grf = Graphics.FromImage(bmp))
        {
            // Fully qualified Aspose.Drawing types to avoid ambiguity with Aspose.Words.
            grf.DrawString(
                error.Message,
                new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, Aspose.Drawing.FontStyle.Regular),
                Aspose.Drawing.Brushes.Red,
                new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
        }
        return bmp;
    }

    public static Stream ConvertImageToWord(Bitmap bmp)
    {
        MemoryStream ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        ms.Position = 0;
        return ms;
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator implementation required for rendered outputs.
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

            if (!parameters.DisplayText)
                gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;
            else
                gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

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
        return GetBarcodeImage(parameters);
    }
}
