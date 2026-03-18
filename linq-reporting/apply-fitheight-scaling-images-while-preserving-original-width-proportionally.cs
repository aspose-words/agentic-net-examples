using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImageFitHeightExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source image file.
            string imagePath = ImageDir + "Logo.jpg";

            // Desired height for the image shape (in points; 1 point = 1/72 inch).
            double fitHeight = 200.0;

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the image as a shape.
            Shape shape = builder.InsertImage(imagePath);

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

            // Save the document.
            doc.Save(ArtifactsDir + "Image.FitHeight.docx");
        }

        // Placeholder properties for the example; replace with actual paths in real code.
        private static string ImageDir => @"C:\Images\";
        private static string ArtifactsDir => @"C:\Output\";
    }
}
