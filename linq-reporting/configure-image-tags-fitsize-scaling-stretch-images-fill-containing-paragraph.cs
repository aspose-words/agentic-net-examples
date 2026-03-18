using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ImageFitSizeExample
{
    static void Main()
    {
        // Path to the folder that contains the source image.
        string imageDir = @"C:\Images\";               // <-- adjust as needed
        string imageFile = Path.Combine(imageDir, "sample.jpg");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will contain the image.
        builder.Writeln("Below is an image that stretches to fill the paragraph width:");

        // Insert the image as an inline shape.
        Shape imageShape = builder.InsertImage(imageFile);

        // Make the image fill the width of the paragraph (page width minus margins).
        double pageWidth = builder.PageSetup.PageWidth;
        double leftMargin = builder.PageSetup.LeftMargin;
        double rightMargin = builder.PageSetup.RightMargin;
        imageShape.Width = pageWidth - leftMargin - rightMargin;

        // Fit the image data to the new shape size, stretching it if necessary.
        imageShape.ImageData.FitImageToShape();

        // Optional: set the height to match the shape's new width while preserving aspect ratio.
        // If you want the image to completely fill the paragraph height as well,
        // you could assign a fixed height here instead.
        // imageShape.Height = imageShape.ImageData.ImageSize.HeightPoints *
        //                     (imageShape.Width / imageShape.ImageData.ImageSize.WidthPoints);

        // Ensure the image is treated as an inline shape (no text wrapping).
        imageShape.WrapType = WrapType.Inline;

        // Save the document to HTML with scaling of images to the shape size.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ScaleImageToShapeSize = true   // Scale images during export so they match the shape dimensions.
        };

        string outputPath = @"C:\Output\ImageFitSize.html";   // <-- adjust as needed
        doc.Save(outputPath, htmlOptions);
    }
}
