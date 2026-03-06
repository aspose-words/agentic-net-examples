using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading; // Added for HtmlLoadOptions

class ExtractImagesFromHtml
{
    static void Main()
    {
        // Path to the HTML file that contains the document.
        string htmlFilePath = @"C:\Docs\Sample.html";

        // Directory where extracted images will be saved.
        string outputImageDir = @"C:\Docs\ExtractedImages\";

        // Base URI for relative image references inside the HTML.
        // This should point to the folder that contains the image files referenced by the HTML.
        string baseUri = @"C:\Docs\Images\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputImageDir);

        // Load the HTML document with a BaseUri so that relative image links can be resolved.
        // Use HtmlLoadOptions (the correct class for HTML loading) instead of the generic LoadOptions.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions
        {
            BaseUri = baseUri
        };
        Document doc = new Document(htmlFilePath, loadOptions);

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

                // Build a unique file name for the extracted image.
                string imageFileName = $"ExtractedImage_{imageIndex}{extension}";
                string imagePath = Path.Combine(outputImageDir, imageFileName);

                // Save the image data to the file system.
                shape.ImageData.Save(imagePath);

                Console.WriteLine($"Saved image {imageIndex}: {imagePath}");
                imageIndex++;
            }
        }

        Console.WriteLine($"Extraction complete. {imageIndex} image(s) saved to '{outputImageDir}'.");
    }
}
