using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source MHTML file.
        string mhtmlPath = @"C:\Docs\SourceDocument.mht";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages";
        Directory.CreateDirectory(outputFolder);

        // Load the MHTML document. LoadOptions can be used if a base URI is required for relative images.
        Document doc = new Document(mhtmlPath);

        // Get all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image format.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name for each extracted image.
                string imageFileName = Path.Combine(
                    outputFolder,
                    $"ExtractedImage_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(imageFileName);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
