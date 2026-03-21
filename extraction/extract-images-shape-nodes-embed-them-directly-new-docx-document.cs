using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractAndEmbedImages
{
    static void Main()
    {
        // Create a source document with an embedded image.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);

        // A 1x1 pixel transparent PNG (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream pngStream = new MemoryStream(pngBytes))
        {
            // Insert the image into the source document as a shape.
            Shape shape = sourceBuilder.InsertImage(pngStream);
            shape.Width = 100;
            shape.Height = 100;
        }

        // Create a new empty document where the extracted images will be embedded.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Get all shape nodes from the source document (including those inside groups).
        NodeCollection shapeNodes = sourceDoc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each shape that actually contains an image.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the image data of the shape into a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset stream position before reading.

                // Insert the image into the destination document.
                Shape insertedShape = destBuilder.InsertImage(imageStream);

                // Preserve the original dimensions (optional).
                insertedShape.Width = shape.Width;
                insertedShape.Height = shape.Height;

                // Add a line break after each image for readability.
                destBuilder.Writeln();
            }
        }

        // Save the new document with all images embedded.
        destDoc.Save("EmbeddedImages.docx");
    }
}
