using System;
using System.IO;
using Aspose.Words;
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
        // 1. Create a simple cover image using Aspose.Drawing and save it.
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(600, 800))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with a solid color
                graphics.Clear(Color.LightBlue);

                // Draw some text on the cover
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                    {
                        graphics.DrawString("Cover Page", font, brush, new PointF(100, 350));
                    }
                }
            }

            // Save the image as JPEG
            bitmap.Save(coverImagePath, ImageFormat.Jpeg);
        }

        // ---------------------------------------------------------------
        // 2. Create a DOCX document, insert the cover image, and add content.
        // ---------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert the cover image at the beginning of the document.
        builder.InsertImage(coverImagePath);

        // Add a page break after the cover to start the main content on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample content to the main part of the document.
        builder.Writeln("This is the main content of the document.");
        builder.Writeln("It follows the custom cover page image.");

        // Save the document as DOCX (bootstrap input file).
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 3. Load the DOCX file and convert it to PDF.
        // ---------------------------------------------------------------
        Document doc = new Document(inputDocxPath);
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // 4. Validate that the PDF was created successfully.
        // ---------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The PDF conversion failed; output file was not created.");
        }

        // Optional cleanup of temporary files (comment out if inspection is needed)
        // File.Delete(coverImagePath);
        // File.Delete(inputDocxPath);
    }
}
