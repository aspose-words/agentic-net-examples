using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading; // Needed for HtmlLoadOptions

namespace ExtractImagesFromHtml
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file.
            string htmlFilePath = @"C:\Input\Document.html";

            // Base folder for relative image URIs referenced in the HTML.
            string baseImageFolder = @"C:\Input\Images";

            // Folder where extracted images will be saved.
            string outputFolder = @"C:\Output\ExtractedImages";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the HTML document with a BaseUri so that relative image references can be resolved.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                BaseUri = baseImageFolder
            };
            Document doc = new Document(htmlFilePath, loadOptions);

            // Get all Shape nodes (including inline and floating images).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                // Process only shapes that actually contain image data.
                if (shape.HasImage)
                {
                    // Determine a suitable file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                    // Build the output file name.
                    string outputFilePath = Path.Combine(
                        outputFolder,
                        $"ExtractedImage_{imageIndex}{extension}");

                    // Save the image bytes to the file system.
                    shape.ImageData.Save(outputFilePath);

                    imageIndex++;
                }
            }

            Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
        }
    }
}
