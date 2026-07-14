using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string imagePath = "cover.png";
        string docxPath = "sample.docx";
        string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple cover image using Aspose.Drawing (no System.Drawing)
        // -----------------------------------------------------------------
        // Create a 600x800 pixel bitmap with a solid color background
        using (Bitmap bitmap = new Bitmap(600, 800))
        {
            // Fill the bitmap with a light blue color
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Optionally draw some text on the cover
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    using (Brush brush = new SolidBrush(Color.DarkBlue))
                    {
                        graphics.DrawString("Cover Page", font, brush, new PointF(100, 350));
                    }
                }
            }

            // Save the bitmap to a PNG file
            bitmap.Save(imagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCX document, insert the cover image, and add content
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the cover image at the beginning of the document
        builder.InsertImage(imagePath);
        // Insert a page break so following content starts on a new page
        builder.InsertBreak(BreakType.PageBreak);

        // Add some sample content after the cover page
        builder.Writeln("This is the main document content.");
        builder.Writeln("Generated on: " + DateTime.Now);

        // Save the document as DOCX (bootstrap input file)
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the DOCX file and convert it to PDF
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 4. Validation – ensure the PDF was created
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF conversion failed: output file not found.");

        // Clean up temporary files (optional)
        File.Delete(imagePath);
        File.Delete(docxPath);
    }
}
