using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample 1x1 PNG image file.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        const string imagePath = "sample.png";
        File.WriteAllBytes(imagePath, imageBytes);

        // -----------------------------------------------------------------
        // Step 1: Create a document that contains a picture content control.
        // -----------------------------------------------------------------
        Document seedDoc = new Document();
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(seedDoc, SdtType.Picture, MarkupLevel.Inline);
        Shape pictureShape = new Shape(seedDoc, ShapeType.Image);
        pictureShape.ImageData.SetImage(imagePath);
        pictureShape.WrapType = WrapType.Inline;
        pictureSdt.AppendChild(pictureShape);
        Paragraph seedParagraph = seedDoc.FirstSection.Body.FirstParagraph;
        seedParagraph.AppendChild(pictureSdt);
        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document and replace picture content controls with inline images.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);
        var pictureSdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Picture)
            .ToList();

        foreach (StructuredDocumentTag sdt in pictureSdts)
        {
            // Find the inner shape that holds the image.
            Shape innerShape = sdt.GetChildNodes(NodeType.Shape, true)
                .OfType<Shape>()
                .FirstOrDefault();

            if (innerShape == null)
                continue;

            // Clone the shape so we can insert it outside the content control.
            Shape newImage = (Shape)innerShape.Clone(true);
            newImage.WrapType = WrapType.Inline;

            // Insert the new image before the content control.
            CompositeNode parent = sdt.ParentNode as CompositeNode;
            if (parent != null)
            {
                parent.InsertBefore(newImage, sdt);
                sdt.Remove();
            }
        }

        // -----------------------------------------------------------------
        // Step 3: Save the modified document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
