using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the document that contains shapes.
        Document doc = new Document("Images.docx");

        // Retrieve all shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Counter for naming extracted images.
        int imageIndex = 0;

        // Iterate through each shape and locate images.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Check if the shape actually contains an image.
            if (shape.HasImage)
            {
                // Determine a suitable file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name for the extracted image.
                string fileName = $"ExtractedImage_{imageIndex}{extension}";

                // Save the image data to the file system.
                shape.ImageData.Save(Path.Combine("ExtractedImages", fileName));

                imageIndex++;
            }
        }

        // Optionally, save the (unchanged) document to a new file.
        doc.Save("ProcessedDocument.docx");
    }
}
