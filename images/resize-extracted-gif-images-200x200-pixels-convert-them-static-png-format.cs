using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ExtractAndResizeGifImages
{
    static void Main()
    {
        // Input Word document that contains images.
        string inputFile = Path.Combine(Environment.CurrentDirectory, "Input.docx");

        // Folder where the converted PNG images will be saved.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "ExtractedImages");
        Directory.CreateDirectory(outputFolder);

        // If the input file does not exist, create a simple document with a GIF image.
        if (!File.Exists(inputFile))
        {
            // Minimal 1x1 transparent GIF (base64 encoded).
            const string gifBase64 = "R0lGODdhAQABAPAAAP///wAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==";
            byte[] gifBytes = Convert.FromBase64String(gifBase64);

            using var ms = new MemoryStream(gifBytes);
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.InsertImage(ms);
            doc.Save(inputFile);
        }

        // Load the document.
        Document docLoaded = new Document(inputFile);

        // Iterate through all shapes in the document.
        NodeCollection shapes = docLoaded.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage)
                continue;

            // Check if the embedded image is a GIF.
            if (shape.ImageData.ImageType == ImageType.Gif)
            {
                // Set the shape size to 200x200 pixels (converted to points: 1 point = 1/72 inch, 96 DPI => 1 pixel = 0.75 point).
                const float pixelsToPoints = 72f / 96f; // 0.75
                shape.Width = 200 * pixelsToPoints;
                shape.Height = 200 * pixelsToPoints;

                // Build the output file name.
                string outputPath = Path.Combine(outputFolder, $"Image_{imageIndex}_200x200.png");

                // Render the shape to a PNG file.
                shape.GetShapeRenderer().Save(outputPath, new ImageSaveOptions(SaveFormat.Png));

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted and resized {imageIndex} GIF image(s) to PNG format.");
    }
}
