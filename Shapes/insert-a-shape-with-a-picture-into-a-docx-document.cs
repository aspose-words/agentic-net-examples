using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the image file to be inserted.
        string imagePath = @"C:\Images\Sample.jpg";

        // Path where the resulting DOCX will be saved.
        string outputPath = @"C:\Output\DocumentWithImage.docx";

        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image; InsertImage returns a Shape representing the picture.
        Shape pictureShape = builder.InsertImage(imagePath);

        // Configure the shape as a floating picture.
        pictureShape.WrapType = WrapType.Square;          // Text wraps around the picture.
        pictureShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        pictureShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        pictureShape.HorizontalAlignment = HorizontalAlignment.Right;
        pictureShape.VerticalAlignment = VerticalAlignment.Top;
        pictureShape.Left = 100;   // Position from the left edge of the page (in points).
        pictureShape.Top = 100;    // Position from the top edge of the page (in points).

        // Save the document in DOCX format.
        doc.Save(outputPath);
    }
}
