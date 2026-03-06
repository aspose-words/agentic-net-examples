using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Insert an inline image from a local file (default size).
        // -----------------------------------------------------------------
        // The image will be inserted at the current cursor position.
        builder.InsertImage(@"Images\SampleImage.jpg");

        // Add a paragraph break after the image.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 2. Insert an inline image from a stream with custom dimensions.
        // -----------------------------------------------------------------
        using (FileStream imageStream = File.OpenRead(@"Images\SampleImage.png"))
        {
            // Insert the image with a width of 200 points and a height of 150 points.
            builder.InsertImage(imageStream, 200.0, 150.0);
        }

        // Add a paragraph break after the image.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 3. Insert a floating image from a byte array.
        // -----------------------------------------------------------------
        // Load the image bytes into a byte array.
        byte[] imageBytes = File.ReadAllBytes(@"Images\SampleImage.gif");

        // Insert the image as a floating shape positioned 100 points from the left
        // and 100 points from the top of the page, with a width of 250 points
        // and a height of 200 points, using a square text wrap.
        builder.InsertImage(
            imageBytes,
            RelativeHorizontalPosition.Page,   // Horizontal reference point.
            100.0,                             // Left offset.
            RelativeVerticalPosition.Page,     // Vertical reference point.
            100.0,                             // Top offset.
            250.0,                             // Width.
            200.0,                             // Height.
            WrapType.Square);                  // Text wrap type.

        // Add a final paragraph break.
        builder.Writeln();

        // -----------------------------------------------------------------
        // 4. Insert a floating image centered on the page (behind text).
        // -----------------------------------------------------------------
        Shape floatingShape = builder.InsertImage(@"Images\SampleImage.jpg");
        floatingShape.WrapType = WrapType.None;                     // No text wrapping.
        floatingShape.BehindText = true;                            // Place behind text.
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.HorizontalAlignment = HorizontalAlignment.Center;
        floatingShape.VerticalAlignment = VerticalAlignment.Center;

        // Save the document to a file.
        doc.Save("ImageDemo.docx");
    }
}
