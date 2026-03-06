// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Math;
using ZXing;                     // NuGet package ZXing.Net for barcode generation

class Program
{
    static void Main()
    {
        // Create a new blank document and associate a DocumentBuilder with it.
        Document doc = new Document();                 // create rule
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert a formatted numbered list.
        // -------------------------------------------------
        builder.ListFormat.ApplyNumberDefault();       // Use default numbered list style
        builder.Writeln("Item 1: Introduction");
        builder.Writeln("Item 2: Details");
        builder.Writeln("Item 3: Conclusion");
        builder.ListFormat.RemoveNumbers();            // End the list formatting

        builder.Writeln(); // Add a blank paragraph between sections

        // -------------------------------------------------
        // 2. Insert an OfficeMath equation (e.g., quadratic formula).
        // -------------------------------------------------
        // Create the equation as a string in Word field format.
        // The EQ field can represent simple equations.
        // For more complex OMath objects you could use the OMath classes,
        // but using a field keeps the example concise.
        builder.InsertField(@"EQ \o(\a\(\b\))", "");   // Insert equation field

        builder.Writeln(); // Separate from next content

        // -------------------------------------------------
        // 3. Insert a SmartArt shape and perform cold rendering.
        // -------------------------------------------------
        // Insert a generic SmartArt shape (the exact type is not critical for the demo).
        // ShapeType.SmartArt is available in Aspose.Words.Drawing.
        Shape smartArt = builder.InsertShape(ShapeType.SmartArt, 300, 200);
        // Force a cold render of the SmartArt drawing. This updates the internal
        // representation without requiring Word to render it on the fly.
        smartArt.UpdateSmartArtDrawing();               // cold rendering

        builder.Writeln(); // Separate from next content

        // -------------------------------------------------
        // 4. Generate a custom barcode image and insert it.
        // -------------------------------------------------
        // Use ZXing to create a Code128 barcode for the sample text.
        var barcodeWriter = new BarcodeWriterPixelData
        {
            Format = BarcodeFormat.CODE_128,
            Options = new ZXing.Common.EncodingOptions
            {
                Height = 80,
                Width = 300,
                Margin = 0
            }
        };
        var pixelData = barcodeWriter.Write("ABC-12345");

        // Convert the pixel data to a System.Drawing.Bitmap.
        using (var bitmap = new Bitmap(pixelData.Width, pixelData.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb))
        {
            var bitmapData = bitmap.LockBits(
                new Rectangle(0, 0, pixelData.Width, pixelData.Height),
                System.Drawing.Imaging.ImageLockMode.WriteOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppRgb);
            try
            {
                // Copy the raw pixel data into the bitmap's buffer.
                System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }

            // Insert the barcode image inline at the current cursor position.
            builder.InsertImage(bitmap, 300, 80);       // insert image with explicit size
        }

        // -------------------------------------------------
        // 5. Save the document to a DOCX file.
        // -------------------------------------------------
        doc.Save("FormattedListWithMathSmartArtBarcode.docx");   // save rule
    }
}
