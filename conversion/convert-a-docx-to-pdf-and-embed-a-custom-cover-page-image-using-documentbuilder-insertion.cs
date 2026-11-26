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
        // Create a simple cover image using Aspose.Drawing
        string coverImagePath = "cover.png";
        using (Bitmap bitmap = new Bitmap(400, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Draw some text on the image
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                {
                    // Use a brush for the text color and a point for the location
                    using (SolidBrush brush = new SolidBrush(Color.DarkBlue))
                    {
                        graphics.DrawString("Cover Page", font, brush, new PointF(50, 80));
                    }
                }
            }
            bitmap.Save(coverImagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document (input)
        string inputDocxPath = "input.docx";
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the main content of the document.");
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Load the DOCX document
        Document doc = new Document(inputDocxPath);

        // Insert the cover image at the beginning of the document
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentStart();
        builder.InsertImage(coverImagePath);
        builder.InsertBreak(BreakType.PageBreak);

        // Convert to PDF
        string outputPdfPath = "output.pdf";
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
