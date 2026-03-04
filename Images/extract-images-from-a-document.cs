using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExtractImagesExample
{
    static void Main()
    {
        // Path to the source Word document.
        string inputFile = @"C:\Docs\Images.docx";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document from the file system.
        Document doc = new Document(inputFile);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        // Iterate through each shape and save the image if the shape contains one.
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string imageFileName = $"Image_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputFolder, imageFileName);

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
