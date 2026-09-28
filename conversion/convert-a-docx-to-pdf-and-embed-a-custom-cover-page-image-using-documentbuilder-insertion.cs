using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare file names
        const string docxPath = "input.docx";
        const string pdfPath = "output.pdf";
        const string coverImagePath = "cover.png";

        // -----------------------------------------------------------------
        // Step 1: Create a simple cover image using Aspose.Drawing
        // -----------------------------------------------------------------
        const int coverWidth = 600;
        const int coverHeight = 800;
        using (Bitmap bitmap = new Bitmap(coverWidth, coverHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with light gray
                graphics.Clear(Color.LightGray);

                // Draw a simple text in the center
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    // Measure string size
                    SizeF textSize = graphics.MeasureString("Cover Page", font);
                    float x = (coverWidth - textSize.Width) / 2;
                    float y = (coverHeight - textSize.Height) / 2;

                    // Draw the text
                    graphics.DrawString("Cover Page", font, Brushes.Black, x, y);
                }
            }

            // Save the bitmap as PNG
            bitmap.Save(coverImagePath, ImageFormat.Png);
        }

        // Verify that the cover image was created
        if (!File.Exists(coverImagePath))
            throw new InvalidOperationException("Cover image was not created.");

        // -----------------------------------------------------------------
        // Step 2: Create a sample DOCX document
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the main content of the document.");
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX was created
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("Input DOCX was not created.");

        // -----------------------------------------------------------------
        // Step 3: Load the DOCX, insert the cover image at the beginning
        // -----------------------------------------------------------------
        Document doc = new Document(docxPath);
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentStart();
        builder.InsertImage(coverImagePath);

        // -----------------------------------------------------------------
        // Step 4: Convert the document (with cover) to PDF
        // -----------------------------------------------------------------
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Output PDF was not created.");

        // Cleanup temporary files (optional)
        File.Delete(docxPath);
        File.Delete(coverImagePath);
    }
}
