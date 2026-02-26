// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class BarcodeDocumentProcessor
{
    // Generates a simple Code‑128 like barcode as a PNG image.
    private static Image GenerateBarcode(string data, int width, int height)
    {
        // Create a bitmap with white background.
        Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.White);

            // Very simple bar pattern: each character becomes a black bar of fixed width.
            // This is only for demonstration; replace with a proper barcode library if needed.
            int barWidth = Math.Max(1, width / (data.Length * 10));
            int x = 0;
            foreach (char c in data)
            {
                // Alternate black/white bars based on character code parity.
                bool isBlack = ((int)c % 2) == 0;
                if (isBlack)
                {
                    g.FillRectangle(Brushes.Black, x, 0, barWidth, height);
                }
                x += barWidth;
            }
        }
        return bmp;
    }

    static void Main()
    {
        // Input and output file paths.
        string inputDocx = @"C:\Docs\InputDocument.docx";
        string outputDocx = @"C:\Docs\OutputDocument.docx";

        // Load the existing DOCX document.
        Document doc = new Document(inputDocx);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert the generated barcode image (PNG) at the end of the doc.
        // -----------------------------------------------------------------
        using (Image barcodeImg = GenerateBarcode("1234567890", 300, 100))
        {
            // Insert the image inline; you can also specify position/size if required.
            builder.MoveToDocumentEnd();
            builder.InsertParagraph();
            builder.InsertImage(barcodeImg);
        }

        // -------------------------------------------------
        // 2. Add a semi‑transparent text watermark.
        // -------------------------------------------------
        // Create a floating shape that will act as the watermark.
        Shape watermark = builder.InsertShape(ShapeType.TextPlainText, 500, 100);
        watermark.TextPath.Text = "CONFIDENTIAL";
        watermark.TextPath.FontFamily = "Arial";
        watermark.TextPath.FontSize = 48;
        watermark.FillColor = Color.LightGray;
        watermark.OutlineColor = Color.LightGray;
        watermark.Rotation = -40; // Diagonal appearance.
        watermark.WrapType = WrapType.None;
        watermark.BehindText = true;
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        watermark.HorizontalAlignment = HorizontalAlignment.Center;
        watermark.VerticalAlignment = VerticalAlignment.Center;
        watermark.Alpha = 0.3; // 30 % opacity.

        // -------------------------------------------------
        // 3. Insert a simple OfficeMath equation.
        // -------------------------------------------------
        // Use an EQ field to represent a math equation (e.g., a+b=c).
        builder.MoveToDocumentEnd();
        builder.InsertParagraph();
        builder.InsertField("EQ \\o(\\a\\b)", ""); // Displays a + b = c.

        // -------------------------------------------------
        // 4. Embed a SmartArt diagram as an OLE object.
        // -------------------------------------------------
        // Assume a PowerPoint file that contains the desired SmartArt.
        string smartArtPptx = @"C:\Docs\SmartArt.pptx";
        if (File.Exists(smartArtPptx))
        {
            builder.MoveToDocumentEnd();
            builder.InsertParagraph();
            // Insert the PPTX as an embedded OLE object (icon optional).
            builder.InsertOleObject(smartArtPptx, false, false, null);
        }

        // -------------------------------------------------
        // 5. Save the modified document.
        // -------------------------------------------------
        doc.Save(outputDocx, SaveFormat.Docx);
    }
}
