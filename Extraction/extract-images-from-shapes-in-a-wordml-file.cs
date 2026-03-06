using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromWordML
{
    static void Main()
    {
        // Load the WORDML (or DOCX) document.
        Document doc = new Document("Input.docx"); // replace with your WORDML file path

        // Ensure the output directory exists.
        string outputDir = "ExtractedImages";
        Directory.CreateDirectory(outputDir);

        // Retrieve all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name for each extracted image.
                string fileName = $"Image_{imageIndex}{extension}";
                string filePath = Path.Combine(outputDir, fileName);

                // Save the image data to the file system.
                shape.ImageData.Save(filePath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to '{outputDir}'.");
    }
}
