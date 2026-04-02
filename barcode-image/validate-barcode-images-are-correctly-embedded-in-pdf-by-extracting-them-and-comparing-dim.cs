using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace BarcodeImageValidation
{
    // Utility class for barcode generation (adapted to use Aspose.Drawing types only)
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

        // Create a simple error image using Aspose.Drawing types
        public static Aspose.Drawing.Bitmap DrawErrorImage(Exception error)
        {
            Aspose.Drawing.Bitmap bmp = new Aspose.Drawing.Bitmap(200, 50);
            using (Aspose.Drawing.Graphics grf = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 8f, Aspose.Drawing.FontStyle.Regular))
                {
                    grf.DrawString(error.Message, font, Aspose.Drawing.Brushes.Red, new Aspose.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
                }
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

    // Custom barcode generator implementing the required interface
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
            return GetBarcodeImage(parameters);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new Word document and register the custom barcode generator
            Document doc = new Document();
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // First barcode – larger scaling factor (200%)
            FieldDisplayBarcode barcode1 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcode1.BarcodeType = "QR";
            barcode1.BarcodeValue = "ABC123";
            barcode1.BackgroundColor = "0xF8BD69";
            barcode1.ForegroundColor = "0xB5413B";
            barcode1.ErrorCorrectionLevel = "3";
            barcode1.ScalingFactor = "200"; // 200%
            barcode1.SymbolHeight = "1000";
            barcode1.SymbolRotation = "0";
            builder.Writeln();

            // Second barcode – default scaling factor (100%)
            FieldDisplayBarcode barcode2 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcode2.BarcodeType = "QR";
            barcode2.BarcodeValue = "DEF456";
            barcode2.BackgroundColor = "0xF8BD69";
            barcode2.ForegroundColor = "0xB5413B";
            barcode2.ErrorCorrectionLevel = "3";
            barcode2.ScalingFactor = "100"; // 100%
            barcode2.SymbolHeight = "1000";
            barcode2.SymbolRotation = "0";
            builder.Writeln();

            // Update fields to ensure barcodes are rendered
            doc.UpdateFields();

            // Save the document as PDF
            string pdfPath = "Barcodes.pdf";
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Load the generated PDF
            Document pdfDoc = new Document(pdfPath);

            // Extract all images (barcode images) from the PDF
            List<(int Width, int Height)> imageDimensions = new List<(int, int)>();
            NodeCollection shapes = pdfDoc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    byte[] imgBytes = shape.ImageData.ImageBytes;
                    using (MemoryStream ms = new MemoryStream(imgBytes))
                    {
                        using (Aspose.Drawing.Bitmap bmp = (Aspose.Drawing.Bitmap)Aspose.Drawing.Image.FromStream(ms))
                        {
                            imageDimensions.Add((bmp.Width, bmp.Height));
                        }
                    }
                }
            }

            // Simple validation: the first barcode (larger scaling) should be wider than the second
            if (imageDimensions.Count >= 2)
            {
                var first = imageDimensions[0];
                var second = imageDimensions[1];
                bool isValid = first.Width > second.Width && first.Height == second.Height;
                Console.WriteLine(isValid
                    ? "Validation succeeded: first barcode image is larger as expected."
                    : "Validation failed: image dimensions do not match expected scaling.");
                Console.WriteLine($"First barcode size: {first.Width}x{first.Height}");
                Console.WriteLine($"Second barcode size: {second.Width}x{second.Height}");
            }
            else
            {
                Console.WriteLine("No barcode images were found in the PDF.");
            }
        }
    }
}
