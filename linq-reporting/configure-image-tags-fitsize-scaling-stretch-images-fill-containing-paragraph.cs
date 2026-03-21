using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class ImageFitSizeExample
{
    static void Main()
    {
        // Create a temporary image file (a tiny red dot PNG).
        string tempDir = Path.GetTempPath();
        string imageFile = Path.Combine(tempDir, "sample.png");
        if (!File.Exists(imageFile))
        {
            // PNG data for a 1x1 red pixel.
            const string base64Png = 
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imageFile, pngBytes);
        }

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

        // Ensure the image is treated as an inline shape (no text wrapping).
        imageShape.WrapType = WrapType.Inline;

        // Save the document to HTML with scaling of images to the shape size.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ScaleImageToShapeSize = true // Scale images during export so they match the shape dimensions.
        };

        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImageFitSize.html");
        doc.Save(outputPath, htmlOptions);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
