using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Input\MapDocument.docx";

        // Folder where extracted PNG images will be saved.
        string outputFolder = @"C:\Output\ExtractedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the Word document.
        Document doc = new Document(sourcePath);

        // Retrieve all shape nodes (including OLE objects, images, etc.).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Check if the shape contains an image (e.g., picture, map object).
            if (shape.HasImage)
            {
                // Prepare high‑resolution PNG save options.
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    // 300 DPI is a common high‑resolution setting.
                    Resolution = 300,
                    // Ensure the background is transparent if the source permits.
                    PaperColor = System.Drawing.Color.Transparent
                };

                // Build a unique file name for each extracted image.
                string pngPath = Path.Combine(outputFolder,
                    $"MapImage_{imageIndex}{FileFormatUtil.ImageTypeToExtension(ImageType.Png)}");

                // Render the shape to a PNG file using the shape renderer.
                shape.GetShapeRenderer().Save(pngPath, pngOptions);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
