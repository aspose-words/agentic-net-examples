using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file to be inserted.
        string imagePath = @"C:\Images\Sample.jpg";

        // Insert the image inline at the current cursor position.
        Shape imageShape = builder.InsertImage(imagePath);

        // Example: set the wrap type (optional, default is inline).
        imageShape.WrapType = WrapType.Inline;

        // Save the document as a DOCX file.
        doc.Save(@"C:\Output\ImageInserted.docx");
    }
}
