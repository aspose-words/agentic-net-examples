using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (JPEG) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.jpg";
        const int imageWidth = 100;
        const int imageHeight = 100;

        // Create a white bitmap and draw a red rectangle.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imageWidth, imageHeight))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
            {
                graphics.DrawRectangle(pen, 10, 10, imageWidth - 20, imageHeight - 20);
            }

            // Save as JPEG – the file extension determines the format.
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document that contains a footnote with the image.
        // -----------------------------------------------------------------
        const string docPath = "FootnoteImages.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Main body text.
        builder.Writeln("This paragraph contains a footnote reference.");

        // Insert a footnote reference.
        builder.InsertFootnote(FootnoteType.Footnote, string.Empty);

        // Retrieve the footnote that was just created.
        Footnote footnote = (Footnote)doc.GetChildNodes(NodeType.Footnote, true).Last();

        // Insert the image into the footnote using a Shape.
        Shape imgShape = new Shape(doc, ShapeType.Image);
        imgShape.ImageData.SetImage(sampleImagePath);
        footnote.FirstParagraph.AppendChild(imgShape);

        // Add some additional text after the image inside the footnote.
        footnote.FirstParagraph.AppendChild(new Run(doc, "Footnote text after the image."));

        // Save the document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract images that reside inside footnotes.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection footnoteNodes = loadedDoc.GetChildNodes(NodeType.Footnote, true);

        int extractedImageCount = 0;
        int footnoteIndex = 0;

        foreach (Footnote fn in footnoteNodes.OfType<Footnote>())
        {
            footnoteIndex++;

            // Find all Shape nodes that contain images inside the current footnote.
            NodeCollection shapeNodes = fn.GetChildNodes(NodeType.Shape, true);
            int imageInFootnoteIndex = 0;

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine the appropriate file extension for the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outputFileName = $"footnote-{footnoteIndex}-{imageInFootnoteIndex}{extension}";

                    // Save the extracted image.
                    shape.ImageData.Save(outputFileName);
                    extractedImageCount++;
                    imageInFootnoteIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImageCount == 0)
            throw new InvalidOperationException("No images were extracted from footnotes.");

        // Optional clean‑up (commented out – uncomment if you want to delete temporary files).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
