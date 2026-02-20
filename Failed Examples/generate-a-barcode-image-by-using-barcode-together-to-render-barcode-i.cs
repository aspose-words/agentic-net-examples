// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

namespace BarcodeExample
{
    // Simple implementation of IBarcodeGenerator that creates a placeholder image.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        // Generates a barcode image based on the supplied parameters.
        // For demonstration purposes this method creates a bitmap with the barcode value drawn as text.
        public Image GetBarcodeImage(BarcodeParameters parameters)
        {
            // Define image size.
            const int width = 300;
            const int height = 100;

            // Create a bitmap and draw the barcode value.
            Bitmap bitmap = new Bitmap(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Font font = new Font("Arial", 24))
                {
                    // Center the text.
                    SizeF textSize = graphics.MeasureString(parameters.BarcodeValue, font);
                    PointF location = new PointF((width - textSize.Width) / 2, (height - textSize.Height) / 2);
                    graphics.DrawString(parameters.BarcodeValue, font, Brushes.Black, location);
                }
            }

            return bitmap;
        }

        // Not used in this example, but required by the interface.
        public Image GetOldBarcodeImage(BarcodeParameters parameters)
        {
            // Reuse the same implementation for old-fashioned barcode fields.
            return GetBarcodeImage(parameters);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Assign the custom barcode generator to the document's field options.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Define barcode parameters (example: QR code).
            BarcodeParameters barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "QR",
                BarcodeValue = "ABC123",
                BackgroundColor = "0xFFFFFF",
                ForegroundColor = "0x000000",
                ErrorCorrectionLevel = "2",
                ScalingFactor = "250",
                SymbolHeight = "1000",
                SymbolRotation = "0"
            };

            // Generate the barcode image using the custom generator.
            Image barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);

            // Insert the generated image into the document.
            builder.InsertImage(barcodeImage);

            // Save the document as PDF using PdfSaveOptions.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            doc.Save("BarcodeDocument.pdf", saveOptions);
        }
    }
}
