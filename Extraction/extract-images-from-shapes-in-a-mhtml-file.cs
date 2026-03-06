using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ExtractImagesFromMhtml
{
    static void Main()
    {
        // Path to the source MHTML file.
        string mhtmlPath = @"C:\Input\document.mht";

        // Directory where extracted images will be saved.
        string outputDir = @"C:\Output\Images";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the MHTML document.
        Document doc = new Document(mhtmlPath);

        // Get all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and extract images.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the full file name for the extracted image.
                string imageFileName = Path.Combine(outputDir, $"Image_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(imageFileName);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputDir}\".");
    }
}
