using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // A 1x1 pixel transparent PNG encoded in Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Desired position and size (in points).
        double leftOffset = 50;   // points from the left margin
        double topOffset = 30;    // points from the top margin
        double imageWidth = 100;  // width in points
        double imageHeight = 50;  // height in points

        // Insert the image as a floating shape with absolute positioning.
        Shape shape = builder.InsertImage(
            imageBytes,
            RelativeHorizontalPosition.Margin, leftOffset,
            RelativeVerticalPosition.Margin, topOffset,
            imageWidth, imageHeight,
            WrapType.None);

        // Place the image behind the text.
        shape.BehindText = true;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderImage.docx");
        doc.Save(outputPath);
    }
}
