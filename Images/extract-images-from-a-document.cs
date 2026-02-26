using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImages
{
    static void Main()
    {
        // Path to the source Word document
        string inputPath = "Images.docx";

        // Directory where extracted images will be saved
        string outputDir = "ExtractedImages";
        Directory.CreateDirectory(outputDir);

        // Load the document
        Document doc = new Document(inputPath);

        // Retrieve all shape nodes (inline and floating)
        var shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine file extension based on the image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = Path.Combine(outputDir, $"Image_{imageIndex}{extension}");

                // Save the image to the file system
                shape.ImageData.Save(fileName);
                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} images to '{outputDir}'.");
    }
}
