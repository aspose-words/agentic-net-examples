using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

namespace BarcodeDisplayExample
{
    // Utility class for barcode image handling and conversions.
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

        public static Bitmap DrawErrorImage(Exception error)
        {
            Bitmap bmp = new Bitmap(200, 100);
            using (Graphics grf = Graphics.FromImage(bmp))
            {
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 10f, Aspose.Drawing.FontStyle.Regular))
                {
                    grf.DrawString(error.Message, font, Aspose.Drawing.Brushes.Red, new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
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

    // Custom barcode generator that Aspose.Words will call for DISPLAYBARCODE fields.
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
            return GetBarcodeImage(parameters);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Register the custom barcode generator so that DISPLAYBARCODE fields are rendered.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Insert a DISPLAYBARCODE field and configure it as a QR code.
            FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            qrField.BarcodeType = "QR";
            qrField.BarcodeValue = "https://example.com";
            qrField.BackgroundColor = "0xFFFFFF"; // white background
            qrField.ForegroundColor = "0x000000"; // black bars
            qrField.ErrorCorrectionLevel = "3";   // highest error correction
            qrField.ScalingFactor = "250";        // 250%
            qrField.SymbolHeight = "1000";        // height in twips
            qrField.SymbolRotation = "0";         // no rotation

            // Add a line break after the field for readability.
            builder.Writeln();

            // Ensure fields are updated so the barcode image is generated.
            doc.UpdateFields();

            // Save the document. The QR code will be visible when the document is opened in Word
            // or when the document is converted to a rendered format (PDF, etc.).
            doc.Save("DisplayBarcodeQRCode.docx");
        }
    }
}
