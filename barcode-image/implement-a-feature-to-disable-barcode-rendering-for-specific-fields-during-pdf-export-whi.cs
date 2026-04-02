using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a barcode that will be rendered.
        builder.Writeln("Rendered barcode:");
        InsertBarcodeField(builder, "CODE128", "123456", displayText: true, disableRendering: false);

        // Insert a barcode that should NOT be rendered in PDF.
        builder.Writeln("\nBarcode disabled for PDF:");
        InsertBarcodeField(builder, "CODE128", "DISABLE:789012", displayText: true, disableRendering: true);

        // Update fields to ensure they are ready.
        doc.UpdateFields();

        // Save as DOCX (both barcodes are present).
        doc.Save("Barcodes.docx");

        // Register custom barcode generator for PDF export.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Save as PDF (the disabled barcode will be blank).
        doc.Save("Barcodes.pdf", SaveFormat.Pdf);
    }

    private static void InsertBarcodeField(DocumentBuilder builder, string barcodeType, string barcodeValue, bool displayText, bool disableRendering)
    {
        // Insert a typed DISPLAYBARCODE field.
        Field field = builder.InsertField(FieldType.FieldDisplayBarcode, true);
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)field;

        // Set barcode properties.
        barcodeField.BarcodeType = barcodeType;
        barcodeField.BarcodeValue = barcodeValue;
        barcodeField.DisplayText = displayText;

        // No additional styling needed for this example.
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator utilities.
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

    public static Aspose.Drawing.Color ConvertColor(string inputColor, Aspose.Drawing.Color defVal)
    {
        if (string.IsNullOrEmpty(inputColor))
            return defVal;
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

    public static Bitmap DrawErrorImage(Exception error)
    {
        Bitmap bmp = new Bitmap(200, 50);
        using (Graphics grf = Graphics.FromImage(bmp))
        {
            grf.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 8f))
            {
                grf.DrawString(error.Message, font, Brushes.Red, new RectangleF(0, 0, bmp.Width, bmp.Height));
            }
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

    public static Stream CreateBlankImage()
    {
        Bitmap bmp = new Bitmap(1, 1);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.Transparent);
        }
        return ConvertImageToWord(bmp);
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator that disables rendering for values prefixed with "DISABLE:".
internal class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // If the barcode value starts with the disable marker, return a blank image.
        if (parameters.BarcodeValue != null && parameters.BarcodeValue.StartsWith("DISABLE:", StringComparison.OrdinalIgnoreCase))
        {
            return CustomBarcodeGeneratorUtils.CreateBlankImage();
        }

        try
        {
            BarcodeGenerator gen = new BarcodeGenerator(
                CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType),
                parameters.BarcodeValue);

            gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.ForegroundColor, gen.Parameters.Barcode.BarColor);
            gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.BackgroundColor, gen.Parameters.BackColor);

            gen.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText ? CodeLocation.Below : CodeLocation.None;

            if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel))
                gen.Parameters.Barcode.QR.ErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(parameters.ErrorCorrectionLevel, gen.Parameters.Barcode.QR.ErrorLevel);

            if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                gen.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(parameters.SymbolRotation, gen.Parameters.RotationAngle);

            double scalingFactor = 1;
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scalingFactor);

            if (gen.BarcodeType == EncodeTypes.QR)
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scalingFactor));
            else
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));

            if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0, Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(parameters.SymbolHeight, gen.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

            if (!string.IsNullOrEmpty(parameters.PosCodeStyle))
                CustomBarcodeGeneratorUtils.SetPosCodeStyle(gen, parameters.PosCodeStyle, parameters.BarcodeValue);

            return CustomBarcodeGeneratorUtils.ConvertImageToWord(gen.GenerateBarCodeImage());
        }
        catch (Exception e)
        {
            return CustomBarcodeGeneratorUtils.ConvertImageToWord(CustomBarcodeGeneratorUtils.DrawErrorImage(e));
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}
