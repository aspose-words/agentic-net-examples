using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace BarcodeHeightWidthExample
{
    internal static class CustomBarcodeGeneratorUtils
    {
        public static double TwipsToPixels(string valueInTwips, double resolution, double defVal)
        {
            try
            {
                int twips = int.Parse(valueInTwips);
                return (twips / 1440.0) * resolution;
            }
            catch
            {
                return defVal;
            }
        }

        public static double TwipsToPixels(string valueInTwips, double defVal)
        {
            // 96 DPI is the default screen resolution used by Aspose.Words when converting twips to pixels.
            return TwipsToPixels(valueInTwips, 96, defVal);
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

        public static QRErrorLevel GetQRCorrectionLevel(string level, QRErrorLevel def)
        {
            return level switch
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
                int argb = Convert.ToInt32(inputColor, 16);
                // Input is assumed to be BGR (as Word does). Convert to ARGB.
                int r = (argb >> 16) & 0xFF;
                int g = (argb >> 8) & 0xFF;
                int b = argb & 0xFF;
                return Aspose.Drawing.Color.FromArgb(r, g, b);
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
            var bmp = new Aspose.Drawing.Bitmap(200, 100);
            using (var grf = Aspose.Drawing.Graphics.FromImage(bmp))
            {
                var font = new Aspose.Drawing.Font("Arial", 8);
                var brush = Aspose.Drawing.Brushes.Red;
                grf.DrawString(error.Message, font, brush, new Aspose.Drawing.RectangleF(0, 0, bmp.Width, bmp.Height));
            }
            return bmp;
        }

        public static Stream ConvertImageToWord(Aspose.Drawing.Bitmap bmp)
        {
            var ms = new MemoryStream();
            bmp.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;
            return ms;
        }
    }

    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
        {
            try
            {
                // Create generator with appropriate symbology.
                BarcodeGenerator gen = new BarcodeGenerator(
                    CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType),
                    parameters.BarcodeValue);

                // Colors.
                gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(
                    parameters.ForegroundColor, gen.Parameters.Barcode.BarColor);
                gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(
                    parameters.BackgroundColor, gen.Parameters.BackColor);

                // Text display.
                gen.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText
                    ? CodeLocation.Below
                    : CodeLocation.None;

                // QR error correction.
                if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel))
                {
                    gen.Parameters.Barcode.QR.ErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(
                        parameters.ErrorCorrectionLevel, gen.Parameters.Barcode.QR.ErrorLevel);
                }

                // Rotation.
                if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                {
                    gen.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(
                        parameters.SymbolRotation, gen.Parameters.RotationAngle);
                }

                // Scaling factor.
                double scalingFactor = 1.0;
                if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                {
                    scalingFactor = CustomBarcodeGeneratorUtils.ScaleFactor(parameters.ScalingFactor, scalingFactor);
                }

                // X dimension (module width) – use default or scaling factor.
                if (gen.BarcodeType == EncodeTypes.QR)
                {
                    gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0,
                        Math.Round(CustomBarcodeGeneratorUtils.DefaultQRXDimensionInPixels * scalingFactor));
                }
                else
                {
                    gen.Parameters.Barcode.XDimension.Pixels = (float)Math.Max(1.0,
                        Math.Round(CustomBarcodeGeneratorUtils.Default1DXDimensionInPixels * scalingFactor));
                }

                // Bar height – use SymbolHeight if supplied.
                if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                {
                    double heightPixels = CustomBarcodeGeneratorUtils.TwipsToPixels(parameters.SymbolHeight,
                        gen.Parameters.Barcode.BarHeight.Pixels);
                    gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0,
                        Math.Round(heightPixels * scalingFactor));
                }

                // Generate image.
                return CustomBarcodeGeneratorUtils.ConvertImageToWord(gen.GenerateBarCodeImage());
            }
            catch (Exception ex)
            {
                // Return an error image so the document still renders.
                return CustomBarcodeGeneratorUtils.ConvertImageToWord(
                    CustomBarcodeGeneratorUtils.DrawErrorImage(ex));
            }
        }

        public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
        {
            // For this example the old API behaves the same as the new one.
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

            // Add a paragraph describing the barcode.
            builder.Writeln("Barcode with custom height (twips):");

            // Insert a DISPLAYBARCODE field using the typed API.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Configure the barcode.
            barcodeField.BarcodeType = "CODE128";
            barcodeField.BarcodeValue = "1234567890";

            // Height is expressed in twips (1 inch = 1440 twips).
            // Here we set it to 720 twips (0.5 inch) to demonstrate the switch.
            barcodeField.SymbolHeight = "720"; // 0.5 inch height.

            // Show the human?readable text below the barcode.
            barcodeField.DisplayText = true;

            // Update fields so that the field properties are stored.
            doc.UpdateFields();

            // Register the custom barcode generator – required for PDF/image output.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Save the document to PDF (rendered barcode image will be embedded).
            string outputPath = "BarcodeOutput.pdf";
            doc.Save(outputPath);
        }
    }
}
