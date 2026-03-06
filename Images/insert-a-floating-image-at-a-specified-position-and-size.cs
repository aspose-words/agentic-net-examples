using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating image.
        // Position: 100 points from the left margin, 150 points from the top margin.
        // Size: 200 points wide, 150 points high.
        // Wrap type: Square (text wraps around the image).
        Shape imageShape = builder.InsertImage(
            "C:\\Images\\Sample.jpg",                     // Path to the image file.
            RelativeHorizontalPosition.Margin, 100,      // Horizontal position.
            RelativeVerticalPosition.Margin, 150,        // Vertical position.
            200,                                          // Width.
            150,                                          // Height.
            WrapType.Square);                            // Text wrap.

        // Additional optional settings.
        imageShape.BehindText = false; // Place the image in front of the text.
        imageShape.WrapType = WrapType.Square;

        // Save the document to disk.
        doc.Save("FloatingImage.docx");
    }
}
