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
        // Paths – adjust as needed.
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";
        const string barcodePngPath = "barcode.png";

        // Load the existing DOCX document (lifecycle rule: use Document ctor with filename).
        Document doc = new Document(inputDocPath);

        // -----------------------------------------------------------------
        // 1. Create a custom barcode image and save it as PNG.
        // -----------------------------------------------------------------
        using (Bitmap bmp = new Bitmap(200, 80))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.White);
            using (Font font = new Font("Arial", 24, FontStyle.Bold))
            {
                // Simple placeholder barcode – replace with a real generator if desired.
                g.DrawString("1234567890", font, Brushes.Black, new PointF(10, 20));
            }
            bmp.Save(barcodePngPath, System.Drawing.Imaging.ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Insert the barcode image into the document.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.InsertParagraph();                     // Ensure we are on a new line.
        builder.InsertImage(barcodePngPath);          // Inline image insertion.

        // -----------------------------------------------------------------
        // 3. Apply the same image as a watermark on every page.
        // -----------------------------------------------------------------
        using (FileStream imgStream = new FileStream(barcodePngPath, FileMode.Open, FileAccess.Read))
        using (Image watermarkImg = Image.FromStream(imgStream))
        {
            doc.Watermark.SetImage(watermarkImg);
        }

        // -----------------------------------------------------------------
        // 4. Insert an OfficeMath equation (simple a+b).
        // -----------------------------------------------------------------
        builder.MoveToDocumentEnd();
        builder.InsertParagraph();
        // EQ field code renders a mathematical equation.
        builder.InsertField("EQ \\o\\(a,b\\)", true);

        // -----------------------------------------------------------------
        // 5. Insert a SmartArt object.
        // -----------------------------------------------------------------
        builder.MoveToDocumentEnd();
        builder.InsertParagraph();

        // Aspose.Words provides a SmartArt shape type. If the specific
        // InsertSmartArt API is unavailable, we can create a generic SmartArt shape.
        Shape smartArt = new Shape(doc, ShapeType.SmartArt);
        smartArt.Width = 300;
        smartArt.Height = 200;
        // Optionally set a layout or style here if needed.
        builder.InsertNode(smartArt);

        // -----------------------------------------------------------------
        // 6. Save the modified document (lifecycle rule: use Document.Save).
        // -----------------------------------------------------------------
        doc.Save(outputDocPath);
    }
}
