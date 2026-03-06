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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Generate a simple barcode image (placeholder) ----------
        using (Bitmap barcodeBmp = new Bitmap(200, 80))
        {
            using (Graphics g = Graphics.FromImage(barcodeBmp))
            {
                g.Clear(Color.White);
                using (Font font = new Font("Arial", 36, FontStyle.Bold))
                {
                    // Draw the barcode text. Replace with a real barcode generator if needed.
                    g.DrawString("1234567890", font, Brushes.Black, new PointF(10, 20));
                }
            }

            // Insert the barcode image as a diagonal watermark.
            Shape watermark = builder.InsertImage(barcodeBmp);
            watermark.WrapType = WrapType.None;                     // No text wrapping.
            watermark.BehindText = true;                           // Place behind document text.
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.Rotation = -45;                              // Diagonal angle.
        }

        // Add a paragraph break after the watermark.
        builder.Writeln();

        // ---------- Insert an OLE object (e.g., an Excel spreadsheet) ----------
        string excelPath = @"C:\Temp\Sample.xlsx"; // Adjust the path to an existing file.
        if (File.Exists(excelPath))
        {
            builder.Writeln("Embedded Excel OLE object:");
            using (FileStream fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                // Insert as an embedded OLE object (not as an icon) using the appropriate ProgID.
                builder.InsertOleObject(fs, "Excel.Sheet", false, null);
            }
        }

        // ---------- Insert an online video ----------
        builder.Writeln();
        builder.Writeln("Online video:");
        string videoUrl = "https://www.youtube.com/watch?v=dQw4w9WgXcQ";
        // Insert a video placeholder with a size of 320x180 points.
        builder.InsertOnlineVideo(videoUrl, 320, 180);

        // Save the document to a DOCX file.
        doc.Save(@"C:\Temp\BarcodeWatermarkOleVideo.docx");
    }
}
