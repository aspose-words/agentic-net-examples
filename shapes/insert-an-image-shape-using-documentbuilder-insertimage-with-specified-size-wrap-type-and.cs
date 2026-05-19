using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class InsertImageShapeExample
{
    public static void Main()
    {
        // Create an output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG image from a Base64 string.
        // This is a 1x1 pixel transparent PNG.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = Path.Combine(outputDir, "sample.png");
        File.WriteAllBytes(imagePath, imageBytes);

        // Verify the image file was created.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException("Failed to create the sample image.", imagePath);

        // Create a new Word document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define positioning and size for the floating image.
        double left = 50; // points from the left margin
        double top = 100; // points from the top margin
        double width = ConvertUtil.PixelToPoint(150); // convert 150 pixels to points
        double height = ConvertUtil.PixelToPoint(150); // convert 150 pixels to points

        // Insert the image as a floating shape with the specified parameters.
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, left,
            RelativeVerticalPosition.Margin, top,
            width, height,
            WrapType.Square);

        // Optional formatting.
        imageShape.BehindText = false;
        imageShape.WrapType = WrapType.Square;

        // Save the document.
        string docPath = Path.Combine(outputDir, "ImageShapeExample.docx");
        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception("The document was not saved correctly.");
    }
}
