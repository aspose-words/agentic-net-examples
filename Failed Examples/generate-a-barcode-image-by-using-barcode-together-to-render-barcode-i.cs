// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode;
using Aspose.BarCode.Generation;

namespace BarcodeExample
{
    // Simple implementation of IBarcodeGenerator that uses Aspose.BarCode to create barcode images.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        // Generates a barcode image based on the supplied parameters and returns it as a stream.
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            // Create a BarCodeBuilder and configure it according to the parameters.
            var builder = new BarCodeBuilder
            {
                // Map the generic BarcodeType to the specific SymbologyType enum.
                // For simplicity, handle a few common types; others can be added as needed.
                SymbologyType = GetSymbology(parameters.BarcodeType),
                CodeText = parameters.BarcodeValue
            };

            // Optional visual settings.
            if (!string.IsNullOrEmpty(parameters.BackgroundColor))
                builder.BackColor = System.Drawing.ColorTranslator.FromHtml(parameters.BackgroundColor);
            if (!string.IsNullOrEmpty(parameters.ForegroundColor))
                builder.ForeColor = System.Drawing.ColorTranslator.FromHtml(parameters.ForegroundColor);
            if (!string.IsNullOrEmpty(parameters.ErrorCorrectionLevel) && parameters.BarcodeType == "QR")
                builder.QRErrorLevel = (QRErrorLevel)Enum.Parse(typeof(QRErrorLevel), parameters.ErrorCorrectionLevel);
            if (!string.IsNullOrEmpty(parameters.ScalingFactor))
                builder.XDimension = Convert.ToInt32(parameters.ScalingFactor) / 100.0f; // approximate scaling
            if (!string.IsNullOrEmpty(parameters.SymbolHeight))
                builder.ImageHeight = Convert.ToInt32(parameters.SymbolHeight) / 1440; // convert twips to points

            // Generate the image into a memory stream.
            var ms = new MemoryStream();
            builder.Save(ms, BarCodeImageFormat.Png);
            ms.Position = 0;
            return ms;
        }

        // Not required for this example; return null.
        public Stream GetOldBarcodeImage(BarcodeParameters parameters) => null;

        // Helper to map string barcode type to SymbologyType.
        private static SymbologyType GetSymbology(string type)
        {
            return type?.ToUpperInvariant() switch
            {
                "QR" => SymbologyType.QR,
                "EAN13" => SymbologyType.EAN13,
                "CODE39" => SymbologyType.Code39Standard,
                "ITF14" => SymbologyType.ITF14,
                _ => SymbologyType.Code128 // default fallback
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Assign the custom barcode generator to the document's field options.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Define barcode parameters for a QR code.
            BarcodeParameters barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "QR",
                BarcodeValue = "ABC123",
                BackgroundColor = "#F8BD69", // using HTML hex notation
                ForegroundColor = "#B5413B",
                ErrorCorrectionLevel = "3",
                ScalingFactor = "250",
                SymbolHeight = "1000",
                SymbolRotation = "0"
            };

            // Generate the barcode image and insert it into the document.
            using (Stream img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
            {
                // Insert the image at the current cursor position.
                builder.InsertImage(img);
            }

            // Save the document as PDF.
            doc.Save("BarcodeOutput.pdf");
        }
    }
}
