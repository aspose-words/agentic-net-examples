using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageExtractor
{
    // Extracts all images from the specified document and saves them to the output folder.
    public static void ExtractImages(string inputFilePath, string outputFolder)
    {
        // Load the document from the input file.
        Document doc = new Document(inputFilePath);

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Retrieve all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Process only shapes that actually contain image data.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the full file name for the extracted image.
                string fileName = Path.Combine(outputFolder, $"Image_{imageIndex}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(fileName);
                imageIndex++;
            }
        }
    }
}

class Program
{
    static void Main()
    {
        // Example usage: specify the source document and the folder for extracted images.
        string sourceDoc = @"C:\Docs\Images.docx";
        string targetFolder = @"C:\Docs\ExtractedImages";

        ImageExtractor.ExtractImages(sourceDoc, targetFolder);
    }
}
