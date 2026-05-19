using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace BarcodeColorExample
{
    // Utility class for barcode parameter conversion.
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
    }

    // Custom barcode generator that Aspose.Words will use when rendering fields.
    internal class CustomBarcodeGenerator : IBarcodeGenerator
    {
        public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
        {
            try
            {
                BarcodeGenerator gen = new BarcodeGenerator(
                    CustomBarcodeGeneratorUtils.GetBarcodeEncodeType(parameters.BarcodeType),
                    parameters.BarcodeValue);

                // Apply colors.
                gen.Parameters.Barcode.BarColor = CustomBarcodeGeneratorUtils.ConvertColor(
                    parameters.ForegroundColor, gen.Parameters.Barcode.BarColor);
                gen.Parameters.BackColor = CustomBarcodeGeneratorUtils.ConvertColor(
                    parameters.BackgroundColor, gen.Parameters.BackColor);

                // Text visibility.
                gen.Parameters.Barcode.CodeTextParameters.Location = parameters.DisplayText
                    ? CodeLocation.Below
                    : CodeLocation.None;

                // QR error correction.
                if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel))
                    gen.Parameters.Barcode.QR.ErrorLevel = CustomBarcodeGeneratorUtils.GetQRCorrectionLevel(
                        parameters.ErrorCorrectionLevel, gen.Parameters.Barcode.QR.ErrorLevel);

                // Rotation.
                if (!string.IsNullOrEmpty(parameters.SymbolRotation))
                    gen.Parameters.RotationAngle = CustomBarcodeGeneratorUtils.GetRotationAngle(
                        parameters.SymbolRotation, gen.Parameters.RotationAngle);

                // Scaling.
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

                // Symbol height.
                if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                    gen.Parameters.Barcode.BarHeight.Pixels = (float)Math.Max(5.0,
                        Math.Round(CustomBarcodeGeneratorUtils.TwipsToPixels(
                            parameters.SymbolHeight, 96, gen.Parameters.Barcode.BarHeight.Pixels) * scalingFactor));

                // POS code style handling.
                if (!string.IsNullOrEmpty(parameters.PosCodeStyle))
                    CustomBarcodeGeneratorUtils.SetPosCodeStyle(gen, parameters.PosCodeStyle, parameters.BarcodeValue);

                // Generate image and return as PNG stream.
                using (var img = gen.GenerateBarCodeImage())
                {
                    MemoryStream ms = new MemoryStream();
                    img.Save(ms, Aspose.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;
                    return ms;
                }
            }
            catch
            {
                // In case of error return an empty stream.
                return new MemoryStream();
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
            // Output PDF file.
            const string outputFile = "BarcodeCustomColors.pdf";

            // Create a new document and register the custom barcode generator.
            Document doc = new Document();
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DISPLAYBARCODE field and customize its colors.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeType = "QR";
            barcodeField.BarcodeValue = "CustomColorTest";
            // Background: light red, Foreground: blue.
            barcodeField.BackgroundColor = "0xFFCCCC";
            barcodeField.ForegroundColor = "0x0000FF";
            // Optional additional settings.
            barcodeField.ErrorCorrectionLevel = "3";
            barcodeField.ScalingFactor = "200";
            barcodeField.SymbolHeight = "800";
            barcodeField.SymbolRotation = "0";

            // Update fields to render the barcode.
            doc.UpdateFields();

            // Save the document as PDF (rendered barcode image will be embedded).
            doc.Save(outputFile);

            // Verify that the document contains at least one shape (the rendered barcode image).
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            Console.WriteLine($"Document saved to '{outputFile}'.");
            Console.WriteLine($"Number of shapes (barcode images) in the document: {shapes.Count}");
        }
    }
}
