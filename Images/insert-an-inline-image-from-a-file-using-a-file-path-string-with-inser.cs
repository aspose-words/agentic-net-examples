using System;
using Aspose.Words;

class InsertImageExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Path to the image file to be inserted.
        string imagePath = @"C:\Images\Sample.jpg";

        // Insert the image inline at the current cursor position.
        builder.InsertImage(imagePath);

        // Save the document to disk.
        doc.Save(@"C:\Output\ImageInserted.docx");
    }
}
