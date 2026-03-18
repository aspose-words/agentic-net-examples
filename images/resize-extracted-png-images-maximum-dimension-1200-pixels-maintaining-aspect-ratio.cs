using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ResizePngImages
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Folder where the resized PNG images will be saved.
        string outputFolder = "ResizedImages";
        Directory.CreateDirectory(outputFolder);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that contain an image.
            if (!shape.HasImage) continue;

            // Obtain the original image dimensions in pixels.
            ImageSize imgSize = shape.ImageData.ImageSize;
            int originalWidth = imgSize.WidthPixels;
            int originalHeight = imgSize.HeightPixels;

            // Determine the larger dimension of the image.
            int maxOriginal = Math.Max(originalWidth, originalHeight);

            // If the image is already within the required size, skip it.
            if (maxOriginal <= 1200) continue;

            // Compute the scaling factor needed to bring the largest side to 1200 pixels.
            float scale = 1200f / maxOriginal;

            // Configure ImageSaveOptions to apply the calculated scale.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Scale = scale // Maintains aspect ratio.
            };

            // Create a unique file name for the resized image.
            string fileName = Path.Combine(outputFolder,
                $"Image_{shape.GetHashCode()}.png");

            // Render the shape (the original image) to a PNG file using the scaling factor.
            shape.GetShapeRenderer().Save(fileName, saveOptions);
        }
    }
}
