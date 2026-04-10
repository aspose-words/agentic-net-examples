using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing namespace for image creation
using Aspose.Drawing.Imaging;      // For ImageFormat
using DrawingFont = Aspose.Drawing.Font;   // Alias to avoid ambiguity with Aspose.Words.Font

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a picture content control (block level) at the current cursor position.
        // Use the documented InsertStructuredDocumentTag method to ensure correct placement.
        StructuredDocumentTag pictureSdt = builder.InsertStructuredDocumentTag(SdtType.Picture);

        // -----------------------------------------------------------------
        // Create a simple PNG image in memory and save it to a temporary file.
        // This file will be referenced by the picture content control.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(Environment.CurrentDirectory, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (DrawingFont font = new DrawingFont("Arial", 12))
                using (SolidBrush brush = new SolidBrush(Color.Black))
                {
                    graphics.DrawString("Img", font, brush, new PointF(10, 40));
                }
            }
            // Save the bitmap to the file system.
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // Create an image shape that loads the image from the file.
        // The SetImage method embeds the image data into the document.
        // -----------------------------------------------------------------
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(imagePath); // Embed the image.
        imageShape.Width = ConvertUtil.PixelToPoint(100);
        imageShape.Height = ConvertUtil.PixelToPoint(100);

        // A picture content control must contain a paragraph (block container).
        Paragraph para = new Paragraph(doc);
        para.AppendChild(imageShape);

        // Add the paragraph (with the image) as a child of the picture content control.
        pictureSdt.AppendChild(para);

        // Save the document; the image is now embedded.
        doc.Save("PictureContentControl.docx");

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }
}
