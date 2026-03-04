using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageExtractor
{
    // Extracts all images from shapes in a DOCX file and saves them to the specified folder.
    public static void ExtractImages(string docxPath, string outputFolder)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(docxPath);

        // Get all shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and process those that contain an image.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string imageFileName = Path.Combine(outputFolder, $"ExtractedImage_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(imageFileName);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }

    // Example usage.
    static void Main()
    {
        string inputDocx = @"C:\Docs\Sample.docx";
        string outputDir = @"C:\Docs\ExtractedImages";

        ExtractImages(inputDocx, outputDir);
    }
}
