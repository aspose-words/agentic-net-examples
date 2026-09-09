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
        // Paths for temporary files
        const string coverImagePath = "cover.png";
        const string docxPath = "sample.docx";
        const string pdfPath = "output.pdf";

        // --------------------------------------------------------------
        // Create a simple cover image using Aspose.Drawing (no System.Drawing)
        // --------------------------------------------------------------
        const int imageWidth = 600;
        const int imageHeight = 800;

        using (Bitmap bitmap = new Bitmap(imageWidth, imageHeight))
        {
            // Obtain a Graphics object for drawing on the bitmap
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background
                graphics.Clear(Color.White);

                // Prepare font and brush
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48);
                try
                {
                    // Draw centered text
                    string text = "Cover Page";
                    // Measure text size
                    SizeF textSize = graphics.MeasureString(text, font);
                    float x = (imageWidth - textSize.Width) / 2;
                    float y = (imageHeight - textSize.Height) / 2;
                    graphics.DrawString(text, font, Brushes.Black, x, y);
                }
                finally
                {
                    font.Dispose();
                }
            }

            // Save the image to a file (PNG format)
            bitmap.Save(coverImagePath, ImageFormat.Png);
        }

        // --------------------------------------------------------------
        // Create a DOCX document, insert the cover image, and add content
        // --------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert the cover image at the beginning
        builder.InsertImage(coverImagePath);
        // Add a page break after the cover
        builder.InsertBreak(BreakType.PageBreak);
        // Add sample body content
        builder.Writeln("This is the main document content after the cover page.");

        // Save the DOCX file
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // --------------------------------------------------------------
        // Load the DOCX and convert it to PDF
        // --------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);

        // --------------------------------------------------------------
        // Validate that the PDF was created
        // --------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional cleanup (commented out to allow inspection of generated files)
        // File.Delete(coverImagePath);
        // File.Delete(docxPath);
    }
}
