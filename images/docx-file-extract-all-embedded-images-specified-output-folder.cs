using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageExtractor
{
    static void Main(string[] args)
    {
        // Determine base directory (where the executable resides).
        string baseDir = AppContext.BaseDirectory;

        // Path to the source DOCX file (relative to the executable).
        string inputPath = Path.Combine(baseDir, "Sample.docx");

        // Folder where extracted images will be saved (relative to the executable).
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // If the sample document does not exist, create a minimal one.
        if (!File.Exists(inputPath))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(inputPath);
        }

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes (including those that contain images).
        var shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Construct a unique file name for each extracted image.
                string fileName = $"Image_{imageIndex}{extension}";

                // Save the image data to the specified folder.
                shape.ImageData.Save(Path.Combine(outputFolder, fileName));

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
