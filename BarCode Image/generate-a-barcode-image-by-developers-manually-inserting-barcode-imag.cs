using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;
#if NET5_0_OR_GREATER
using SkiaSharp;
#endif

// Simple custom barcode generator that returns a placeholder image as a stream.
// In a real scenario you would generate a proper barcode based on the parameters.
public class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Helper that creates a dummy PNG image and returns it as a MemoryStream.
    private static Stream CreatePlaceholderImage(int width = 200, int height = 200)
    {
#if NET5_0_OR_GREATER
        using var bitmap = new SKBitmap(width, height);
        using var canvas = new SKCanvas(bitmap);
        // Fill background with a light colour.
        canvas.Clear(new SKColor(0xF8, 0xBD, 0x69));
        // Draw a simple rectangle to visualise the placeholder.
        var paint = new SKPaint
        {
            Color = new SKColor(0xB5, 0x41, 0x3B),
            Style = SKPaintStyle.Stroke,
            StrokeWidth = 4
        };
        canvas.DrawRect(new SKRect(2, 2, width - 2, height - 2), paint);
        // Encode to PNG.
        using var image = SKImage.FromBitmap(bitmap);
        var data = image.Encode(SKEncodedImageFormat.Png, 100);
        var ms = new MemoryStream();
        data.SaveTo(ms);
        ms.Position = 0;
        return ms;
#else
        // Fallback – return an empty stream if SkiaSharp is not available.
        return new MemoryStream();
#endif
    }

    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        // Here you could inspect "parameters" and generate a real barcode.
        // For the purpose of the example we just return a placeholder image.
        return CreatePlaceholderImage();
    }

    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        // Same placeholder for the legacy field type.
        return CreatePlaceholderImage();
    }
}

class Program
{
    static void Main()
    {
        // Create a new blank document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Assign the custom barcode generator to the document's field options.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // ---------- QR CODE ----------
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

        // Generate the barcode image as a stream.
        using var qrStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
        // Optionally save the image to disk for verification.
        using (var file = File.Create("FieldOptions.BarcodeGenerator.QR.png"))
        {
            qrStream.CopyTo(file);
        }
        // Reset the stream position before inserting.
        qrStream.Position = 0;
        builder.InsertImage(qrStream);

        // ---------- EAN13 BARCODE ----------
        barcodeParameters = new BarcodeParameters
        {
            BarcodeType = "EAN13",
            BarcodeValue = "501234567890",
            DisplayText = true,
            PosCodeStyle = "CASE",
            FixCheckDigit = true
        };
        using var ean13Stream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
        using (var file = File.Create("FieldOptions.BarcodeGenerator.EAN13.png"))
        {
            ean13Stream.CopyTo(file);
        }
        ean13Stream.Position = 0;
        builder.InsertImage(ean13Stream);

        // ---------- CODE39 BARCODE ----------
        barcodeParameters = new BarcodeParameters
        {
            BarcodeType = "CODE39",
            BarcodeValue = "12345ABCDE",
            AddStartStopChar = true
        };
        using var code39Stream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
        using (var file = File.Create("FieldOptions.BarcodeGenerator.CODE39.png"))
        {
            code39Stream.CopyTo(file);
        }
        code39Stream.Position = 0;
        builder.InsertImage(code39Stream);

        // ---------- ITF14 BARCODE ----------
        barcodeParameters = new BarcodeParameters
        {
            BarcodeType = "ITF14",
            BarcodeValue = "09312345678907",
            CaseCodeStyle = "STD"
        };
        using var itf14Stream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
        using (var file = File.Create("FieldOptions.BarcodeGenerator.ITF14.png"))
        {
            itf14Stream.CopyTo(file);
        }
        itf14Stream.Position = 0;
        builder.InsertImage(itf14Stream);

        // Save the resulting DOCX document.
        doc.Save("FieldOptions.BarcodeGenerator.docx", SaveFormat.Docx);
    }
}
