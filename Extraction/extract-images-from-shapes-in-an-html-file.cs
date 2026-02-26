using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the HTML file that contains the shapes with images.
        string htmlFilePath = @"C:\Input\Document.html";

        // Base folder for relative image URIs referenced inside the HTML.
        string baseImageFolder = @"C:\Input\Images";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Output\ExtractedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the HTML document. HtmlLoadOptions.BaseUri allows Aspose.Words to resolve
        // relative image paths that are referenced in the HTML file.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            BaseUri = baseImageFolder
        };
        Document doc = new Document(htmlFilePath, loadOptions);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine a suitable file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string imageFileName = $"ExtractedImage_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputFolder, imageFileName);

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                Console.WriteLine($"Saved image {imageIndex}: {imagePath}");
                imageIndex++;
            }
        }

        Console.WriteLine($"Extraction complete. {imageIndex} image(s) saved to \"{outputFolder}\".");
    }
}
