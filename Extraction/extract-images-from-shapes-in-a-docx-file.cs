using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        // Iterate through each shape and extract images.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"ExtractedImage_{imageIndex}{extension}";
                string outputPath = Path.Combine("ExtractedImages", fileName);

                // Ensure the output directory exists.
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // Save the image data to the file system.
                shape.ImageData.Save(outputPath);
                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to the 'ExtractedImages' folder.");
    }
}
