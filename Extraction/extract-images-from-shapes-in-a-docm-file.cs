using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExtractImagesFromDocm
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourcePath = @"C:\Input\DocumentWithImages.docm";

        // Directory where extracted images will be saved.
        string outputDir = @"C:\Output\ExtractedImages\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the DOCM document.
        Document doc = new Document(sourcePath);

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract the image if present.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string imageFileName = $"Image_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputDir}\".");
    }
}
