using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string coverImagePath = "cover.jpg";
        const string inputDocxPath = "input.docx";
        const string outputPdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple cover page image using Aspose.Drawing (no System.Drawing)
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(600, 800))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background
                graphics.Clear(Color.LightBlue);

                // Prepare font and brush
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48);
                Brush brush = Brushes.Black;

                // Draw text centered
                string text = "Cover Page";
                SizeF textSize = graphics.MeasureString(text, font);
                float x = (bitmap.Width - textSize.Width) / 2;
                float y = (bitmap.Height - textSize.Height) / 2;
                graphics.DrawString(text, font, brush, new PointF(x, y));
            }

            // Save the image as JPEG
            bitmap.Save(coverImagePath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Create a sample DOCX file (input document)
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the main document content.");
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the DOCX, insert the cover image at the beginning, and convert to PDF
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocxPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move cursor to the very start of the document
        builder.MoveToDocumentStart();

        // Insert the cover image
        builder.InsertImage(coverImagePath);

        // Optional: add a page break after the cover page
        builder.InsertBreak(BreakType.PageBreak);

        // Save the final document as PDF
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 4. Validation: ensure the PDF was created
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Cleanup temporary files (optional)
        try { File.Delete(coverImagePath); } catch { }
        try { File.Delete(inputDocxPath); } catch { }

        // Indicate success (no console output required per requirements)
    }
}
