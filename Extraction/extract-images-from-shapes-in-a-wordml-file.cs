using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the WORDML file.
        string dataDir = @"C:\Data\";

        // Load the WORDML document.
        Document doc = new Document(Path.Combine(dataDir, "WordMLFile.xml"));

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract the image if it has one.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image format.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string outputPath = Path.Combine(dataDir, $"ExtractedImage_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(outputPath);

                imageIndex++;
            }
        }
    }
}
