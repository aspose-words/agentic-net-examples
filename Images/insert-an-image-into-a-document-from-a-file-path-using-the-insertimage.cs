using System;
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
        string imagePath = @"C:\Images\Sample.jpg";

        // Insert the image inline at the current cursor position.
        Shape imageShape = builder.InsertImage(imagePath);

        // Example: modify the inserted image's properties if needed.
        // imageShape.WrapType = WrapType.Square;

        // Save the document with the inserted image.
        doc.Save(@"C:\Output\ImageInserted.docx");
    }
}
