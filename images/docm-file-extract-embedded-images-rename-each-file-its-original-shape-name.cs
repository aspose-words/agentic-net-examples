using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExtractAndRenameImages
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourceDocPath = @"C:\Input\Sample.docm";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Output\Images";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCM document (uses the provided Document(string) constructor rule).
        Document doc = new Document(sourceDocPath);

        // Get all shape nodes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true)
                        .OfType<Shape>()
                        .Where(s => s.HasImage);

        foreach (Shape shape in shapes)
        {
            // Determine a safe file name based on the shape's original name.
            // If the shape has no name, fall back to a generated one.
            string baseName = !string.IsNullOrEmpty(shape.Name) ? shape.Name : "UnnamedShape";

            // Remove any characters that are invalid in file names.
            foreach (char c in Path.GetInvalidFileNameChars())
                baseName = baseName.Replace(c, '_');

            // Determine the appropriate file extension for the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

            // Combine folder, base name and extension to create the full path.
            string outputPath = Path.Combine(outputFolder, $"{baseName}{extension}");

            // Save the image data to the file (uses the ImageData.Save(string) method rule).
            shape.ImageData.Save(outputPath);
        }
    }
}
