using Aspose.Words;
using Aspose.Words.Drawing;
using System;

class InsertImageExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify the full path to the image file.
        string imagePath = @"C:\Images\Sample.jpg";

        // Insert the image inline at the current cursor position.
        // The overload InsertImage(string) inserts the image at 100% scale.
        Shape imageShape = builder.InsertImage(imagePath);

        // Optional: modify the inserted shape (e.g., set wrap type, position, etc.).
        // imageShape.WrapType = WrapType.Square;

        // Save the resulting document.
        doc.Save(@"C:\Output\ImageInserted.docx");
    }
}
