using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a simple PNG image programmatically.
        const string imageFileName = "sample.png";
        CreateSampleImage(imageFileName);

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add introductory text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is a picture content control:");

        // Create a block‑level picture content control.
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Block);
        pictureSdt.Title = "SamplePicture";
        pictureSdt.Tag = "sample-picture";

        // Append the content control to the document body.
        doc.FirstSection.Body.AppendChild(pictureSdt);

        // The picture content control must contain a paragraph that holds the image.
        Paragraph pictureParagraph = new Paragraph(doc);
        pictureSdt.AppendChild(pictureParagraph);

        // Insert the image into the paragraph.
        DocumentBuilder imageBuilder = new DocumentBuilder(doc);
        imageBuilder.MoveTo(pictureParagraph);
        Shape pictureShape = imageBuilder.InsertImage(imageFileName);
        pictureShape.WrapType = WrapType.Inline; // Ensure the picture stays inline within the control.

        // Save the resulting document. The image will be embedded automatically.
        const string outputFileName = "picture-content-control.docx";
        doc.Save(outputFileName);

        // Clean up the temporary image file.
        if (File.Exists(imageFileName))
        {
            File.Delete(imageFileName);
        }
    }

    private static void CreateSampleImage(string path)
    {
        // Create a 100x100 pixel bitmap with a simple drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.DarkBlue, 3))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
