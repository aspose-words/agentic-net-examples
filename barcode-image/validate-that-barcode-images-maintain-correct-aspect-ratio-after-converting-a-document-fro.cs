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
        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a typed barcode field.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE128";
        barcodeField.BarcodeValue = "1234567890";
        barcodeField.DisplayText = true;
        barcodeField.ForegroundColor = "000000"; // Black
        barcodeField.BackgroundColor = "FFFFFF"; // White
        barcodeField.ScalingFactor = "100"; // 100%
        barcodeField.SymbolHeight = "720"; // 0.5 inch (720 twips)

        // Register the custom barcode generator.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document as PDF (in memory).
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, SaveFormat.Pdf);
            // PDF saved – in a real scenario we could write it to disk or further process it.
        }

        // Prepare barcode parameters matching the field settings.
        Aspose.Words.Fields.BarcodeParameters parameters = new Aspose.Words.Fields.BarcodeParameters
        {
            BarcodeType = barcodeField.BarcodeType,
            BarcodeValue = barcodeField.BarcodeValue,
            DisplayText = barcodeField.DisplayText,
            ForegroundColor = barcodeField.ForegroundColor,
            BackgroundColor = barcodeField.BackgroundColor,
            ScalingFactor = barcodeField.ScalingFactor,
            SymbolHeight = barcodeField.SymbolHeight,
            SymbolRotation = barcodeField.SymbolRotation,
            ErrorCorrectionLevel = barcodeField.ErrorCorrectionLevel,
            PosCodeStyle = barcodeField.PosCodeStyle
        };

        // Generate the barcode image using the custom generator.
        CustomBarcodeGenerator generator = new CustomBarcodeGenerator();
        using (Stream imageStream = generator.GetBarcodeImage(parameters))
        {
            // Load the image with Aspose.Drawing.
            using (Image img = Image.FromStream(imageStream))
            {
                double aspectRatio = (double)img.Width / img.Height;
                Console.WriteLine($"Generated barcode image size: {img.Width}x{img.Height}");
                Console.WriteLine($"Aspect ratio (width/height): {aspectRatio:F2}");

                // Simple validation: aspect ratio should be greater than 1 for CODE128.
                if (aspectRatio > 1.0)
                {
                    Console.WriteLine("Aspect ratio validation passed.");
                }
                else
                {
                    Console.WriteLine("Aspect ratio validation failed.");
                }
            }
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
        Bitmap bmp = new Bitmap(200, 100);
        using (Graphics grf = Graphics.FromImage(bmp))
        {
            using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 10))
                {
                    grf.DrawString(error.Message, font, brush, new RectangleF(0, 0, bmp.Width, bmp.Height));
                }
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
}

// -----------------------------------------------------------------------------
// Custom barcode generator implementation.
internal class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
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

            double scalingFactor = 1.0;
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scalingFactor);

            if (gen.BarcodeType == EncodeTypes.QR)
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scalingFactor));
            else
                gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0, Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));

            if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0,
                    Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(parameters.SymbolHeight, 96, gen.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

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
