using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image using Aspose.Drawing.
        string imagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.LightBlue);

                // Draw a rectangle.
                using (Pen pen = new Pen(Color.DarkBlue, 2))
                {
                    graphics.DrawRectangle(pen, new Rectangle(10, 10, 80, 80));
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // Create a sample DOCX document and insert the image.
        string inputDocxPath = "input.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content with an embedded image:");
        builder.InsertImage(imagePath);
        sourceDoc.Save(inputDocxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputDocxPath);

        // Save the document as MHTML. Images and fonts are embedded by default.
        string outputMhtmlPath = "output.mhtml";
        doc.Save(outputMhtmlPath, SaveFormat.Mhtml);

        // Validate that the MHTML file was created.
        if (!File.Exists(outputMhtmlPath))
        {
            throw new InvalidOperationException("Expected output MHTML was not created.");
        }

        // Clean up temporary files (optional).
        if (File.Exists(imagePath))
        {
            File.Delete(imagePath);
        }

        if (File.Exists(inputDocxPath))
        {
            File.Delete(inputDocxPath);
        }
    }
}
