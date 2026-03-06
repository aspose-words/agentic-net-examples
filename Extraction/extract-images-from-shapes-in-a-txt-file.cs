using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromTxt
{
    static void Main()
    {
        // Path to the input TXT file. Aspose.Words can load plain text as a document.
        string inputFilePath = @"C:\Input\sample.txt";

        // Directory where extracted images will be saved.
        string outputDirectory = @"C:\Output\Images";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDirectory);

        // Load the TXT file into an Aspose.Words Document.
        Document doc = new Document(inputFilePath);

        // Retrieve all Shape nodes from the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract images.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the full path for the extracted image file.
                string imagePath = Path.Combine(outputDirectory,
                    $"ExtractedImage_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extraction complete. {imageIndex} image(s) saved to {outputDirectory}.");
    }
}
