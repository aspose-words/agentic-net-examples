using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Rendering;
using Aspose.Words.Drawing;

class ResizeExtractedPngImages
{
    static void Main()
    {
        // Load the source Word document that contains PNG images.
        Document doc = new Document(@"C:\Input\DocumentWithImages.docx");

        // Prepare a folder for the resized preview images.
        string previewFolder = @"C:\Output\Previews\";
        Directory.CreateDirectory(previewFolder);

        // Iterate over all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage) continue;

            // Ensure the image is a PNG; skip other formats.
            string imageExtension = shape.ImageData.ImageType.ToString().ToLower();
            if (imageExtension != "png") continue;

            // Build a file name for the preview image.
            string previewPath = Path.Combine(previewFolder, $"ImagePreview_{imageIndex++}.png");

            // Configure ImageSaveOptions to render the shape at 50% scale.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Scale = 0.5 reduces both width and height by half.
                Scale = 0.5f
            };

            // Render the shape to a PNG file using the configured options.
            // ShapeRenderer inherits from NodeRendererBase and provides a Save method.
            shape.GetShapeRenderer().Save(previewPath, saveOptions);
        }

        // Optionally, inform that processing is complete.
        Console.WriteLine("Preview images have been generated.");
    }
}
