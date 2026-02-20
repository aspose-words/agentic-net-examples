using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using SkiaSharp;

namespace BarcodeDemo
{
    // Custom barcode generator that creates a simple placeholder image using SkiaSharp.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        // Generates a barcode image and returns it as a Stream.
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            // Create a bitmap of fixed size.
            const int width = 200;
            const int height = 200;
            using var bitmap = new SKBitmap(width, height);
            using var canvas = new SKCanvas(bitmap);

            // Fill background.
            canvas.Clear(SKColors.White);

            // Text to draw – the barcode value.
            string text = parameters?.BarcodeValue ?? "N/A";

            // Prepare paint for drawing the text.
            using var paint = new SKPaint
            {
                Color = SKColors.Black,
                TextSize = 24,
                IsAntialias = true,
                TextAlign = SKTextAlign.Center
            };

            // Calculate position so the text is centred.
            float x = width / 2f;
            // SKCanvas draws text from the baseline, so we adjust using FontMetrics.
            var fm = paint.FontMetrics;
            float y = height / 2f - (fm.Ascent + fm.Descent) / 2f;

            canvas.DrawText(text, x, y, paint);

            // Encode the bitmap to PNG and return a MemoryStream.
            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            var stream = new MemoryStream();
            data.SaveTo(stream);
            stream.Position = 0; // Reset position for the caller.
            return stream;
        }

        // For legacy BARCODE fields – reuse the same implementation.
        public Stream GetOldBarcodeImage(BarcodeParameters parameters)
        {
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

            // Assign the custom barcode generator to the document.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Define barcode parameters (example: QR code).
            BarcodeParameters barcodeParameters = new BarcodeParameters
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

            // Generate the barcode image as a stream.
            using Stream barcodeStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);

            // Insert the generated image into the document.
            builder.InsertImage(barcodeStream);

            // Save the document in DOCX format.
            doc.Save("CustomBarcode.docx");
        }
    }
}
