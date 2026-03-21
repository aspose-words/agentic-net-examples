using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageFitHeightExample
{
    class Program
    {
        static void Main()
        {
            // Create a tiny PNG image (1x1 pixel) in memory.
            byte[] imageBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
            using var imageStream = new MemoryStream(imageBytes);

            // Desired height for the image shape (in points; 1 point = 1/72 inch).
            double fitHeight = 200.0;

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the image as a shape from the memory stream.
            Shape shape = builder.InsertImage(imageStream);

            // Retrieve the original image dimensions (in points).
            ImageSize originalSize = shape.ImageData.ImageSize;

            // Calculate the width that preserves the original aspect ratio.
            double aspectRatio = originalSize.WidthPoints / originalSize.HeightPoints;
            double fitWidth = fitHeight * aspectRatio;

            // Apply the calculated dimensions to the shape.
            shape.Height = fitHeight;
            shape.Width = fitWidth;

            // Ensure the image data fits the new shape dimensions.
            shape.ImageData.FitImageToShape();

            // Ensure the output directory exists.
            Directory.CreateDirectory(ArtifactsDir);

            // Save the document.
            doc.Save(Path.Combine(ArtifactsDir, "Image.FitHeight.docx"));
        }

        // Output directory for the generated document.
        private static string ArtifactsDir => Path.Combine(Path.GetTempPath(), "AsposeOutput");
    }
}
