// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Create a new blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Insert an OfficeMath equation (a over b).
        // -------------------------------------------------
        // Using a field with the EQ code as a quick way to add a simple fraction.
        builder.InsertField("EQ \\o(\\(a\\),\\(b\\))");
        builder.Writeln();

        // -------------------------------------------------
        // 2. Insert a SmartArt diagram and force cold rendering.
        // -------------------------------------------------
        // Insert a SmartArt shape (size 300x200 points).
        Shape smartArt = builder.InsertShape(ShapeType.SmartArt, 300, 200);
        // Update the SmartArt drawing using the cold‑rendering engine.
        smartArt.UpdateSmartArtDrawing();

        builder.Writeln();

        // -------------------------------------------------
        // 3. Add a text watermark.
        // -------------------------------------------------
        doc.Watermark.SetText("CONFIDENTIAL");

        // -------------------------------------------------
        // 4. Generate a simple barcode image and insert it.
        // -------------------------------------------------
        // Create a bitmap that mimics a Code‑128 barcode.
        using (Bitmap bmp = new Bitmap(200, 80))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                string data = "1234567890";
                int x = 10;
                foreach (char c in data)
                {
                    // Alternate bar width for demonstration purposes.
                    int barWidth = (c % 2 == 0) ? 4 : 2;
                    g.FillRectangle(Brushes.Black, x, 10, barWidth, 60);
                    x += barWidth + 2;
                }
            }

            // Save the bitmap to a memory stream and insert it.
            using (MemoryStream ms = new MemoryStream())
            {
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                builder.InsertImage(ms, 200, 80);
            }
        }

        // -------------------------------------------------
        // 5. Save the document.
        // -------------------------------------------------
        doc.Save("OfficeMathSmartArtWatermarkBarcode.docx");
    }
}
