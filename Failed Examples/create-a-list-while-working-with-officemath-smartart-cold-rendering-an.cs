// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Create a simple numbered list ----------
        builder.Writeln("1. First item");
        builder.Writeln("2. Second item");
        builder.Writeln("3. Third item");
        builder.Writeln(); // Add an empty paragraph after the list.

        // ---------- Insert OfficeMath (placeholder) ----------
        // In a real scenario you would create an OfficeMath node (e.g., from MathML) and insert it:
        // OfficeMath officeMath = new OfficeMath(doc);
        // builder.InsertNode(officeMath);
        // For this example we leave a comment indicating where the OfficeMath would go.

        // ---------- Insert SmartArt and trigger cold rendering ----------
        // Insert a SmartArt shape. ShapeType.SmartArt is the enum value for a SmartArt container.
        Shape smartArt = builder.InsertShape(ShapeType.SmartArt, 300, 200);
        // Populate the SmartArt with a layout if needed (omitted for brevity).
        // Force the cold rendering engine to generate the drawing.
        smartArt.UpdateSmartArtDrawing();

        // Add a paragraph break after the SmartArt.
        builder.Writeln();

        // ---------- Generate a custom barcode image and insert it ----------
        // Here we draw a very simple Code128‑like barcode using System.Drawing.
        // In production you would use a dedicated barcode library.
        using (Bitmap bitmap = new Bitmap(200, 80))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);

                // Draw dummy bars.
                for (int i = 0; i < 20; i++)
                {
                    bool isBlack = (i % 2 == 0);
                    Brush brush = isBlack ? Brushes.Black : Brushes.White;
                    graphics.FillRectangle(brush, i * 10, 0, 8, 80);
                }

                // Draw the human‑readable text below the bars.
                using (Font font = new Font("Arial", 12))
                {
                    graphics.DrawString("1234567890", font, Brushes.Black, new PointF(10, 60));
                }
            }

            // Convert the bitmap to a PNG byte array.
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                byte[] imageBytes = ms.ToArray();

                // Insert the barcode image inline, scaling it to the desired size.
                builder.InsertImage(imageBytes, 200, 80);
            }
        }

        // ---------- Save the document ----------
        doc.Save("OfficeMathSmartArtBarcode.docx");
    }
}
