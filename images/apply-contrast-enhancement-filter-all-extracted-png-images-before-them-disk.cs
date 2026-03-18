using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Ensure the output directory exists.
        Directory.CreateDirectory("Output");

        // Retrieve all shapes that contain images.
        var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
                             .Cast<Shape>()
                             .Where(s => s.HasImage);

        int imageIndex = 0;
        foreach (var shape in imageShapes)
        {
            // Process only PNG images.
            if (shape.ImageData.ImageType == ImageType.Png)
            {
                // Apply a contrast enhancement.
                // Value must be between 0.0 (least contrast) and 1.0 (greatest contrast).
                shape.ImageData.Contrast = 0.8; // example: increase contrast

                // Build a file name with the correct extension.
                string fileName = $"ExtractedImage_{imageIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";

                // Save the modified image to disk.
                shape.ImageData.Save(Path.Combine("Output", fileName));

                imageIndex++;
            }
        }
    }
}
