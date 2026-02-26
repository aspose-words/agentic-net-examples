// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class DocumentManipulation
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Insert a paragraph that will contain an OfficeMath object.
        //    (Actual OfficeMath creation is not covered by the available API,
        //     so we place a placeholder text that can be replaced later.)
        // ------------------------------------------------------------
        builder.Writeln("Here is an OfficeMath equation placeholder:");
        builder.Writeln("[OfficeMath]");

        // ------------------------------------------------------------
        // 2. Insert a SmartArt shape and perform cold rendering.
        //    The InsertShape method can insert a SmartArt placeholder (ShapeType.Group)
        //    and then we call UpdateSmartArtDrawing to render it.
        // ------------------------------------------------------------
        // Insert a free‑floating shape that will act as a SmartArt container.
        Shape smartArt = builder.InsertShape(
            ShapeType.Group,                     // Using Group as a generic container.
            RelativeHorizontalPosition.Page,     // Position relative to the page.
            100,                                 // Left offset (points).
            RelativeVerticalPosition.Page,
            100,                                 // Top offset (points).
            300,                                 // Width (points).
            200,                                 // Height (points).
            WrapType.None);                      // No text wrapping.

        // Optionally set a title for the SmartArt (visible in Word UI).
        smartArt.Title = "Sample SmartArt";

        // Perform cold rendering – updates the pre‑rendered drawing.
        smartArt.UpdateSmartArtDrawing();

        // Add a paragraph break after the SmartArt.
        builder.Writeln();

        // ------------------------------------------------------------
        // 3. Add a watermark to the document.
        //    The watermark is a semi‑transparent shape placed behind the text.
        // ------------------------------------------------------------
        // Insert a shape that covers most of the page.
        Shape watermark = builder.InsertShape(
            ShapeType.TextPlainText,             // Simple text shape.
            RelativeHorizontalPosition.Page,
            0,
            RelativeVerticalPosition.Page,
            0,
            doc.FirstSection.PageSetup.PageWidth,
            doc.FirstSection.PageSetup.PageHeight,
            WrapType.None);

        // Set the watermark text.
        watermark.TextPath.Text = "CONFIDENTIAL";

        // Rotate the watermark for a typical diagonal appearance.
        watermark.Rotation = -40;

        // Make the shape appear behind the document text.
        watermark.WrapType = WrapType.None;
        watermark.BehindText = true;

        // Set a light gray color with some transparency.
        watermark.FillColor = Color.FromArgb(50, Color.LightGray);

        // Ensure the watermark does not interfere with editing.
        watermark.IsLayoutInCell = false;

        // Move the cursor after the watermark so further content is added normally.
        builder.MoveToDocumentEnd();

        // ------------------------------------------------------------
        // 4. Generate a custom barcode image and insert it.
        //    This example draws a simple Code‑128‑like barcode using System.Drawing.
        // ------------------------------------------------------------
        // Create a bitmap for the barcode.
        using (Bitmap barcodeBmp = new Bitmap(300, 100))
        {
            using (Graphics g = Graphics.FromImage(barcodeBmp))
            {
                g.Clear(Color.White);

                // Simple pattern: alternating black and white bars.
                int barWidth = 4;
                for (int i = 0; i < 50; i++)
                {
                    if (i % 2 == 0)
                    {
                        g.FillRectangle(Brushes.Black, i * barWidth, 0, barWidth, barcodeBmp.Height);
                    }
                }

                // Draw the human‑readable text below the bars.
                using (Font font = new Font("Arial", 12, FontStyle.Bold))
                {
                    string barcodeText = "1234567890";
                    SizeF textSize = g.MeasureString(barcodeText, font);
                    g.DrawString(barcodeText, font, Brushes.Black,
                        (barcodeBmp.Width - textSize.Width) / 2,
                        barcodeBmp.Height - textSize.Height - 5);
                }
            }

            // Save the bitmap to a memory stream.
            using (MemoryStream ms = new MemoryStream())
            {
                barcodeBmp.Save(ms, ImageFormat.Png);
                ms.Position = 0; // Reset stream position before reading.

                // Insert the barcode image inline at the current cursor position.
                builder.InsertImage(ms);
            }
        }

        // ------------------------------------------------------------
        // 5. Save the document to a DOCX file.
        // ------------------------------------------------------------
        doc.Save("ManipulatedDocument.docx", SaveFormat.Docx);
    }
}
