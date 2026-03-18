using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

class RemovePictureContentControls
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Get all StructuredDocumentTag (content control) nodes in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate backwards so that removing nodes does not affect the loop.
        for (int i = sdtNodes.Count - 1; i >= 0; i--)
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)sdtNodes[i];

            // Process only picture content controls.
            if (sdt.SdtType != SdtType.Picture)
                continue;

            // Find the image shape that resides inside the picture content control.
            Shape pictureShape = sdt.GetChildNodes(NodeType.Shape, true)
                                   .Cast<Shape>()
                                   .FirstOrDefault(sh => sh.HasImage);

            if (pictureShape == null)
                continue; // No image found – skip this control.

            // Retrieve the image bytes from the original shape.
            byte[] imageBytes = pictureShape.ImageData.ImageBytes;

            // Create a new inline image shape that will replace the content control.
            Shape inlineShape = new Shape(doc, ShapeType.Image);
            // Use the Stream overload of SetImage – works with the SkiaSharp build of Aspose.Words.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                inlineShape.ImageData.SetImage(ms);
            }
            inlineShape.WrapType = WrapType.Inline;

            // Preserve the original dimensions (optional).
            inlineShape.Width = pictureShape.Width;
            inlineShape.Height = pictureShape.Height;

            // Insert the new inline image after the content control and then remove the control.
            CompositeNode parent = (CompositeNode)sdt.ParentNode;
            parent.InsertAfter(inlineShape, sdt);
            sdt.Remove();
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
