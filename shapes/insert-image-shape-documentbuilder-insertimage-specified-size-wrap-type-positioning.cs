using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file to be inserted.
        string imagePath = @"C:\Images\Sample.png";

        // Insert a floating image with custom size, position and wrap type.
        // - Position is relative to the page margins.
        // - Left offset: 100 points, Top offset: 50 points.
        // - Width: 200 points, Height: 150 points.
        // - Text will wrap around the image in a square fashion.
        Shape imageShape = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.Margin, 100,
            RelativeVerticalPosition.Margin, 50,
            200, 150,
            WrapType.Square);

        // Example of an additional property: keep the image in front of the text.
        imageShape.BehindText = false;

        // Save the document to disk.
        doc.Save(@"C:\Output\ImageInserted.docx");
    }
}
