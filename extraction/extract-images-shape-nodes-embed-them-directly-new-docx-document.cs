using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractAndEmbedImages
{
    static void Main()
    {
        // Paths to the source and destination documents.
        string sourcePath = @"Images.docx";      // Replace with your source file.
        string destinationPath = @"EmbeddedImages.docx";

        // Load the source document that contains shapes with images.
        Document sourceDoc = new Document(sourcePath);

        // Create a new empty document where the extracted images will be embedded.
        Document destDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(destDoc);

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
                // InsertImage returns a Shape representing the inserted picture.
                Shape insertedShape = builder.InsertImage(imageStream);

                // Preserve the original dimensions (optional).
                insertedShape.Width = shape.Width;
                insertedShape.Height = shape.Height;

                // Add a line break after each image for readability.
                builder.Writeln();
            }
        }

        // Save the new document with all images embedded.
        destDoc.Save(destinationPath);
    }
}
