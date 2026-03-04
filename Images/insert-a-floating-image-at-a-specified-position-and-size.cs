using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertFloatingImageExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file to be inserted.
        string imagePath = @"C:\Images\SampleImage.jpg";

        // Desired position (in points) relative to the page margins.
        double left = 100;   // distance from the left margin
        double top = 150;    // distance from the top margin

        // Desired size of the image (in points).
        double width = 200;
        double height = 120;

        // Insert the image as a floating shape with the specified position, size and wrap type.
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, left,
            RelativeVerticalPosition.Margin, top,
            width, height,
            WrapType.Square);

        // Optional: adjust additional layout properties.
        imageShape.BehindText = false;                     // place image in front of text
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;

        // Save the document to a file.
        doc.Save(@"C:\Output\FloatingImage.docx");
    }
}
