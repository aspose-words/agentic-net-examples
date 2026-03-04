using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Retrieve a live collection of all Shape nodes in the document (deep traversal).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each Shape and locate those that contain an image.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage) // Image present in this shape.
            {
                // Output basic information about the found image.
                Console.WriteLine($"Image #{imageIndex}: Type = {shape.ImageData.ImageType}");

                // Optional: save the image to the file system.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"ExtractedImage_{imageIndex}{extension}";
                string outputPath = Path.Combine("OutputImages", fileName);
                Directory.CreateDirectory("OutputImages");
                shape.ImageData.Save(outputPath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Total images found: {imageIndex}");

        // Save the (potentially modified) document if further processing is required.
        doc.Save("Output.docx");
    }
}
