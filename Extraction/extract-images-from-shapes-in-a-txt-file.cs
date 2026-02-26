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
        // Path to the input TXT file. Aspose.Words can load plain text as a document.
        string inputPath = @"C:\Input\sample.txt";

        // Directory where extracted images will be saved.
        string outputDir = @"C:\Output\Images\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the TXT file into a Document object.
        Document doc = new Document(inputPath);

        // Get all Shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract the image if it has one.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string imageFileName = $"ExtractedImage_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputDir}\".");
    }
}
