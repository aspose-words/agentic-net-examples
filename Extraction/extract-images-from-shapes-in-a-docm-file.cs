using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromDocm
{
    static void Main()
    {
        // Load the DOCM file.
        Document doc = new Document("Input.docm");

        // Retrieve all shape nodes in the document (including those inside groups).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Process only shapes that actually contain image data.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"ExtractedImage_{imageIndex}{extension}";

                // Save the image to the file system.
                shape.ImageData.Save(fileName);
                imageIndex++;
            }
        }
    }
}
