using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.BarCode.Generation;
using Aspose.Drawing; // Only for type references; avoid ambiguous Font by using fully qualified names where needed.

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Register the custom barcode generator to ensure proper rendering.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a QR code DISPLAYBARCODE field.
        FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "QR12345";
        qrField.BackgroundColor = "0xFFFFFF";
        qrField.ForegroundColor = "0x000000";
        qrField.ErrorCorrectionLevel = "3";
        qrField.ScalingFactor = "200";
        qrField.SymbolHeight = "800";
        qrField.SymbolRotation = "0";

        builder.Writeln();

        // Insert an ITF14 barcode DISPLAYBARCODE field.
        FieldDisplayBarcode itf14Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        itf14Field.BarcodeType = "ITF14";
        itf14Field.BarcodeValue = "09312345678907";
        itf14Field.CaseCodeStyle = "STD";

        builder.Writeln();

        // Insert an EAN13 barcode DISPLAYBARCODE field.
        FieldDisplayBarcode ean13Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        ean13Field.BarcodeType = "EAN13";
        ean13Field.BarcodeValue = "501234567890";
        ean13Field.DisplayText = true;
        ean13Field.PosCodeStyle = "CASE";
        ean13Field.FixCheckDigit = true;

        builder.Writeln();

        // Insert a CODE39 barcode DISPLAYBARCODE field.
        FieldDisplayBarcode code39Field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "CODE39DATA";
        code39Field.AddStartStopChar = true;

        // Update fields to generate barcode images.
        doc.UpdateFields();

        // Simple verification: output each field's type and field code.
        foreach (Field field in doc.Range.Fields)
        {
            Console.WriteLine($"Field Type: {field.Type}, Code: {field.GetFieldCode()}");
        }

        // Save the document as PDF to force rendering of barcodes.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Barcodes.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Also save as DOCX for reference.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Barcodes.docx");
        doc.Save(docxPath);
    }
}

// Utility class for barcode generation helpers.
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
            // Use fully qualified Font and FontStyle to avoid ambiguity.
            using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, Aspose.Drawing.FontStyle.Regular))
            {
                grf.DrawString(error.Message, font, Aspose.Drawing.Brushes.Red,
                    new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
            }
        }
        return bmp;
    }

    public static Stream ConvertImageToWord(Aspose.Drawing.Bitmap bmp)
    {
        MemoryStream ms = new MemoryStream();
        bmp.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
        ms.Position = 0;
        return ms;
    }
}

// Custom barcode generator implementing the Aspose.Words IBarcodeGenerator interface.
internal class CustomBarcodeGenerator : Aspose.Words.Fields.IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        try
        {
            BarcodeGenerator gen = new BarcodeGenerator(CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType), parameters.BarcodeValue);

            gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.ForegroundColor, gen.Parameters.Barcode.BarColor);
            gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(parameters.BackgroundColor, gen.Parameters.BackColor);

            if (!parameters.DisplayText)
                gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.None;
            else
                gen.Parameters.Barcode.CodeTextParameters.Location = CodeLocation.Below;

            gen.Parameters.Barcode.QR.ErrorLevel = QRErrorLevel.LevelH;
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
        // Delegate to the main method for legacy support.
        return GetBarcodeImage(parameters);
    }
}
