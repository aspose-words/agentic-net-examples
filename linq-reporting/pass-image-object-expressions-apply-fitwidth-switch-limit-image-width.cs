using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class FitWidthExample
{
    static void Main()
    {
        // Paths to the input image and the output document.
        string imagePath = @"C:\Images\Sample.jpg";
        string outputPath = @"C:\Output\FitWidthDocument.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image directly from the file path. This overload returns the Shape that contains the image.
        Shape imageShape = builder.InsertImage(imagePath);

        // Define the maximum width (in points). 1 point = 1/72 inch.
        double maxWidthPoints = 200.0; // Adjust as needed.

        // If the image is wider than the maximum, limit its width.
        if (imageShape.Width > maxWidthPoints)
        {
            // Set the shape's width to the maximum value.
            imageShape.Width = maxWidthPoints;

            // Fit the image data to the new shape size while preserving aspect ratio.
            imageShape.ImageData.FitImageToShape();
        }

        // Save the document to the specified file.
        doc.Save(outputPath);
    }
}
