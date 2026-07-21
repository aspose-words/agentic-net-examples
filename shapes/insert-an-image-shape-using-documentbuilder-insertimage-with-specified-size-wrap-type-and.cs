using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ImageShapeExample
{
    public static void Main()
    {
        // Create a temporary image file programmatically.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "tempImage.png");
        CreateSampleImage(imagePath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a floating shape with specific size, position and wrap type.
        // Parameters: file name, horizontal position reference, left offset,
        // vertical position reference, top offset, width, height, wrap type.
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, 50,   // 50 points from the left margin
            RelativeVerticalPosition.Margin, 50,     // 50 points from the top margin
            150,                                      // width in points
            100,                                      // height in points
            WrapType.Square);                         // text will wrap square around the image

        // Verify that the shape has the expected dimensions.
        if (imageShape.Width != 150 || imageShape.Height != 100)
            throw new InvalidOperationException("Image shape dimensions are not as expected.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImageShapeExample.docx");
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    // Generates a simple PNG image from a Base64 string and writes it to the specified path.
    private static void CreateSampleImage(string path)
    {
        // This is a 1x1 pixel blue PNG image.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcV8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, imageBytes);
    }
}
