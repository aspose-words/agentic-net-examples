using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and register the custom barcode generator.
        Document doc = new Document();
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert various DISPLAYBARCODE fields with predefined switches.
        InsertDisplayBarcode(builder,
            barcodeType: "QR",
            barcodeValue: "ABC123",
            backgroundColor: "0xF8BD69",
            foregroundColor: "0xB5413B",
            errorCorrectionLevel: "3",
            scalingFactor: "250",
            symbolHeight: "1000",
            symbolRotation: "0");

        builder.Writeln();

        InsertDisplayBarcode(builder,
            barcodeType: "EAN13",
            barcodeValue: "501234567890",
            displayText: true,
            posCodeStyle: "CASE",
            fixCheckDigit: true);

        builder.Writeln();

        InsertDisplayBarcode(builder,
            barcodeType: "CODE39",
            barcodeValue: "12345ABCDE",
            addStartStopChar: true);

        builder.Writeln();

        InsertDisplayBarcode(builder,
            barcodeType: "ITF14",
            barcodeValue: "09312345678907",
            caseCodeStyle: "STD");

        // Update fields and save the document.
        doc.UpdateFields();
        doc.Save(Path.Combine(outputDir, "DisplayBarcodes.docx"));
    }

    private static void InsertDisplayBarcode(DocumentBuilder builder,
        string barcodeType,
        string barcodeValue,
        string backgroundColor = null,
        string foregroundColor = null,
        string errorCorrectionLevel = null,
        string scalingFactor = null,
        string symbolHeight = null,
        string symbolRotation = null,
        bool? displayText = null,
        string posCodeStyle = null,
        bool? fixCheckDigit = null,
        bool? addStartStopChar = null,
        string caseCodeStyle = null)
    {
        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set mandatory properties.
        field.BarcodeType = barcodeType;
        field.BarcodeValue = barcodeValue;

        // Set optional switches if they are provided.
        if (!string.IsNullOrEmpty(backgroundColor))
            field.BackgroundColor = backgroundColor;
        if (!string.IsNullOrEmpty(foregroundColor))
            field.ForegroundColor = foregroundColor;
        if (!string.IsNullOrEmpty(errorCorrectionLevel))
            field.ErrorCorrectionLevel = errorCorrectionLevel;
        if (!string.IsNullOrEmpty(scalingFactor))
            field.ScalingFactor = scalingFactor;
        if (!string.IsNullOrEmpty(symbolHeight))
            field.SymbolHeight = symbolHeight;
        if (!string.IsNullOrEmpty(symbolRotation))
            field.SymbolRotation = symbolRotation;
        if (displayText.HasValue)
            field.DisplayText = displayText.Value;
        if (!string.IsNullOrEmpty(posCodeStyle))
            field.PosCodeStyle = posCodeStyle;
        if (fixCheckDigit.HasValue)
            field.FixCheckDigit = fixCheckDigit.Value;
        if (addStartStopChar.HasValue)
            field.AddStartStopChar = addStartStopChar.Value;
        if (!string.IsNullOrEmpty(caseCodeStyle))
            field.CaseCodeStyle = caseCodeStyle;
    }
}

// -----------------------------------------------------------------------------
// Custom barcode generator utilities (as provided in the reference implementation)
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

    public static Aspose.Drawing.Bitmap DrawErrorImage(Exception error)
    {
        Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(100, 100);
        using (Aspose.Drawing.Graphics grf = Aspose.Drawing.Graphics.FromImage(bmp))
        {
            grf.DrawString(error.Message,
                new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, FontStyle.Regular),
                Aspose.Drawing.Brushes.Red,
                new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
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
// Custom barcode generator implementation (as provided in the reference)
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
