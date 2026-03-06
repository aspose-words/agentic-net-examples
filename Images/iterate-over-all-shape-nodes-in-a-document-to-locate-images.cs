using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Linq;

class Program
{
    static void Main()
    {
        // Load the Word document.
        Document doc = new Document("Input.docx");

        // Retrieve all Shape nodes in the document (deep traversal).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each Shape and check if it contains an image.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Output basic information about the image.
                Console.WriteLine($"Image #{imageIndex}: Type={shape.ImageData.ImageType}, Size={shape.ImageData.ImageBytes?.Length ?? 0} bytes");

                // Save the image to the file system (optional).
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"ExtractedImage_{imageIndex}{extension}";
                shape.ImageData.Save(fileName);

                imageIndex++;
            }
        }

        Console.WriteLine($"Total images found: {imageIndex}");
    }
}
