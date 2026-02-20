// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BarcodeDocumentGenerator
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Generate a simple barcode image in memory.
        // -------------------------------------------------
        // For demonstration, we draw the barcode text using a monospaced font.
        // In a real scenario, replace this with a proper barcode generation library.
        const string barcodeText = "123456789012";
        const int barcodeWidth = 300;
        const int barcodeHeight = 100;

        using (Bitmap bitmap = new Bitmap(barcodeWidth, barcodeHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                using (Font font = new Font("Free 3 of 9", 48, FontStyle.Regular, GraphicsUnit.Point))
                {
                    // Draw the barcode text centered.
                    SizeF textSize = graphics.MeasureString("*" + barcodeText + "*", font);
                    PointF location = new PointF(
                        (barcodeWidth - textSize.Width) / 2,
                        (barcodeHeight - textSize.Height) / 2);
                    graphics.DrawString("*" + barcodeText + "*", font, Brushes.Black, location);
                }
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0; // Reset stream position for reading.

                // -------------------------------------------------
                // 2. Insert the barcode image into the document.
                // -------------------------------------------------
                builder.Writeln("Below is the generated barcode:");
                builder.InsertImage(imageStream);
                builder.Writeln(); // Add an empty line after the image.
            }
        }

        // -------------------------------------------------
        // 3. Add a text watermark to the document.
        // -------------------------------------------------
        // The watermark will appear on every page.
        doc.Watermark.SetText("CONFIDENTIAL", new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 72,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        });

        // -------------------------------------------------
        // 4. (Optional) Insert a SmartArt placeholder.
        // -------------------------------------------------
        // Aspose.Words does not expose a direct API for SmartArt creation.
        // As a placeholder, we insert a simple shape that can later be replaced with SmartArt.
        builder.Writeln("SmartArt placeholder:");
        Shape smartArtPlaceholder = new Shape(doc, ShapeType.Rectangle);
        smartArtPlaceholder.Width = 300;
        smartArtPlaceholder.Height = 150;
        smartArtPlaceholder.WrapType = WrapType.Inline;
        smartArtPlaceholder.StrokeColor = Color.Blue;
        smartArtPlaceholder.FillColor = Color.LightBlue;
        builder.InsertNode(smartArtPlaceholder);
        builder.Writeln();

        // -------------------------------------------------
        // 5. Save the document using cold rendering options.
        // -------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            // Cold rendering: disable high‑quality rendering to speed up the process.
            UseHighQualityRendering = false,
            // Ensure anti‑aliasing is also disabled for faster rendering.
            UseAntiAliasing = false
        };

        // Save the document to a file.
        doc.Save("BarcodeDocument.docx", saveOptions);
    }
}
