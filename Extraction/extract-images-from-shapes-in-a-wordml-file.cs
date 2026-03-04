using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the WORDML (WordprocessingML) file.
        string inputPath = @"C:\Docs\SampleWordML.xml";

        // Load the document. Aspose.Words automatically detects the WORDML format.
        Document doc = new Document(inputPath);

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract the image if it has one.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name for the extracted image.
                string outputFile = Path.Combine(
                    @"C:\ExtractedImages",
                    $"ExtractedImage_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(outputFile);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to C:\\ExtractedImages");
    }
}
