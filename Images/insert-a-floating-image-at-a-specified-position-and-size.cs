using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image that will be inserted.
        string imagePath = @"C:\Images\Sample.jpg";

        // Insert a floating image.
        // Horizontal position: 100 points from the left margin.
        // Vertical position:   150 points from the top margin.
        // Size:                200 points wide, 150 points high.
        // Wrap type:           Square (text wraps around the image).
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, 100,
            RelativeVerticalPosition.Margin, 150,
            200, 150,
            WrapType.Square);

        // Example of additional settings (optional).
        imageShape.BehindText = false; // place the image in front of the text.

        // Save the document to a file.
        doc.Save(@"C:\Output\FloatingImage.docx");
    }
}
