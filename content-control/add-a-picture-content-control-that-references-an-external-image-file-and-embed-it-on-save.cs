using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image using Aspose.Drawing.
        string imagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                g.DrawEllipse(new Pen(Color.DarkBlue, 3), 10, 10, 80, 80);
            }
            bitmap.Save(imagePath);
        }

        // Create a new blank document.
        Document doc = new Document();

        // Create a paragraph that will host the picture content control.
        Paragraph hostParagraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(hostParagraph);

        // Create an inline picture content control.
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
        {
            Title = "SampleImage",
            Tag = "sample-image"
        };

        // Insert the picture content control into the paragraph.
        hostParagraph.AppendChild(pictureSdt);

        // Insert the external image into the picture content control.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(pictureSdt);
        builder.InsertImage(imagePath);

        // Save the document; the image will be embedded inside the content control.
        string outputPath = Path.Combine(outputDir, "picture-content-control.docx");
        doc.Save(outputPath);
    }
}
