using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ExtractAndResizeGifImages
{
    static void Main()
    {
        // Input Word document that contains images.
        string inputFile = @"C:\Docs\Input.docx";

        // Folder where the converted PNG images will be saved.
        string outputFolder = @"C:\Docs\ExtractedImages\";
        Directory.CreateDirectory(outputFolder);

        // Load the document (lifecycle rule: load).
        Document doc = new Document(inputFile);

        // Iterate through all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage)
                continue;

            // Check if the embedded image is a GIF.
            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                // Prepare save options for PNG output with the required size (200x200 pixels).
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
                {
                    // The ImageSize property forces the rendered image to the specified pixel dimensions.
                    ImageSize = new Size(200, 200)
                };

                // Build the output file name.
                string outputPath = Path.Combine(outputFolder,
                    $"Image_{imageIndex}_200x200.png");

                // Render the shape to a PNG file using the specified options (lifecycle rule: save).
                shape.GetShapeRenderer().Save(outputPath, saveOptions);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted and resized {imageIndex} GIF image(s) to PNG format.");
    }
}
