using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source MHTML file.
        string mhtmlPath = @"C:\Input\document.mht";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Output\Images\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the MHTML document.
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFormat = LoadFormat.Mhtml
        };
        Document doc = new Document(mhtmlPath, loadOptions);

        // Get all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that contain an image.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"Image_{imageIndex}{extension}";
                string fullPath = Path.Combine(outputFolder, fileName);

                // Save the image data to the file system.
                shape.ImageData.Save(fullPath);

                Console.WriteLine($"Saved image #{imageIndex} to: {fullPath}");
                imageIndex++;
            }
        }

        Console.WriteLine($"Extraction complete. {imageIndex} image(s) saved.");
    }
}
