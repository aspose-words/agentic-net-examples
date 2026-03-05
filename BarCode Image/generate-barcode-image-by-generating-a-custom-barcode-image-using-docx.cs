using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using SkiaSharp;

// Custom barcode generator that creates a simple placeholder image using SkiaSharp.
// This avoids the System.Drawing dependency which is not supported on all platforms.
class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        const int width = 200;
        const int height = 100;

        // Create a bitmap surface.
        using var bitmap = new SKBitmap(width, height);
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);

        // Prepare paint for drawing the barcode value as plain text.
        var paint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 24,
            IsAntialias = true,
            Typeface = SKTypeface.FromFamilyName("Arial")
        };

        string text = parameters?.BarcodeValue ?? string.Empty;
        // Draw the text near the centre of the image.
        canvas.DrawText(text, 10, height / 2 + paint.TextSize / 2, paint);

        // Encode the bitmap to PNG and return it as a stream.
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        var stream = new MemoryStream();
        data.SaveTo(stream);
        stream.Position = 0;
        return stream;
    }

    // For compatibility with older BARCODE fields; reuse the same implementation.
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
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Assign the custom barcode generator to the document.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Set up barcode parameters (example: QR code with value "ABC123").
        var barcodeParameters = new BarcodeParameters
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

        // Generate the barcode image and insert it into the document.
        using (Stream imageStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
        {
            builder.InsertImage(imageStream);
        }

        // Save the document containing the custom barcode image.
        doc.Save("CustomBarcode.docx");
    }
}
