using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromTxt
{
    static void Main()
    {
        // Load the TXT file as a Word document.
        // Aspose.Words can import plain text files.
        Document doc = new Document("Input.txt");

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Counter for naming extracted image files.
        int imageIndex = 0;

        // Iterate through each shape and extract the image if the shape contains one.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string outputFileName = $"ExtractedImage_{imageIndex}{extension}";

                // Save the image data to the file system.
                shape.ImageData.Save(Path.Combine("ExtractedImages", outputFileName));

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to the 'ExtractedImages' folder.");
    }
}
