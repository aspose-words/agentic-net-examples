// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing; // For Image and Bitmap
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace BarcodeToPdfExample
{
    // Simple implementation of IBarcodeGenerator that returns a placeholder image.
    public class SimpleBarcodeGenerator : IBarcodeGenerator
    {
        // Generates a dummy barcode image based on the supplied parameters.
        public Image GetBarcodeImage(BarcodeParameters parameters)
        {
            // Create a simple bitmap (e.g., 200x100) with a solid background.
            // In a real scenario you would generate a proper barcode here.
            Bitmap bitmap = new Bitmap(200, 100);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw the barcode value as text for demonstration purposes.
                using (Font font = new Font("Arial", 12))
                using (Brush brush = new SolidBrush(Color.Black))
                {
                    g.DrawString(parameters.BarcodeValue ?? "N/A", font, brush, new PointF(10, 40));
                }
            }
            return bitmap;
        }

        // Generates a dummy old‑fashioned barcode image (not used in this example).
        public Image GetOldBarcodeImage(BarcodeParameters parameters)
        {
            // Reuse the same placeholder logic.
            return GetBarcodeImage(parameters);
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
            doc.FieldOptions.BarcodeGenerator = new SimpleBarcodeGenerator();

            // Example: generate a QR code barcode.
            BarcodeParameters qrParams = new BarcodeParameters
            {
                BarcodeType = "QR",
                BarcodeValue = "ABC123",
                BackgroundColor = "0xF8BD69",
                ForegroundColor = "0xB5413B",
                ErrorCorrectionLevel = "3",
                ScalingFactor = "250",
                SymbolHeight = "1000",
                SymbolRotation = "0"
            };

            // Obtain the barcode image from the generator.
            Image qrImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(qrParams);

            // Insert the barcode image into the document.
            builder.InsertImage(qrImage);

            // Add a paragraph break for readability.
            builder.Writeln();

            // Example: generate an EAN13 barcode.
            BarcodeParameters eanParams = new BarcodeParameters
            {
                BarcodeType = "EAN13",
                BarcodeValue = "501234567890",
                DisplayText = true,
                PosCodeStyle = "CASE",
                FixCheckDigit = true
            };

            Image eanImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(eanParams);
            builder.InsertImage(eanImage);

            // Save the document as PDF using PdfSaveOptions.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example option: enable high‑quality rendering.
                UseHighQualityRendering = true
            };
            doc.Save("BarcodesOutput.pdf", pdfOptions);
        }
    }
}
