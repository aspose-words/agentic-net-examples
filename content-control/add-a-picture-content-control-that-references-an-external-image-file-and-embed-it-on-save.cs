using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a simple PNG image using Aspose.Drawing.
        const string imageFileName = "sample.png";
        using (Bitmap bitmap = new Bitmap(200, 100))
        {
            // Obtain a Graphics object for the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Use Aspose.Drawing.Font to avoid ambiguity with Aspose.Words.Font.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20, FontStyle.Bold))
                {
                    graphics.DrawString("Aspose", font, Brushes.DarkBlue, new PointF(10, 30));
                }
            }

            bitmap.Save(imageFileName, ImageFormat.Png);
        }

        // Create a new Word document.
        Document doc = new Document();

        // Create a paragraph to hold the content control.
        Paragraph paragraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Create a picture content control (inline).
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
        {
            Title = "SampleImage",
            Tag = "sample-image"
        };

        // Create a shape (image) and load the external image file.
        Shape pictureShape = new Shape(doc, ShapeType.Image);
        pictureShape.ImageData.SetImage(imageFileName); // Embeds the image data.

        // Add the picture shape to the content control.
        pictureSdt.AppendChild(pictureShape);

        // Add the content control to the paragraph.
        paragraph.AppendChild(pictureSdt);

        // Save the document; the image will be embedded.
        const string outputFileName = "PictureContentControl.docx";
        doc.Save(outputFileName);
    }
}
