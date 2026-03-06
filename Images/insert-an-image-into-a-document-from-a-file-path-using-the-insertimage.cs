using System;
using Aspose.Words;

class InsertImageExample
{
    static void Main()
    {
        // Path to the image file to be inserted
        string imagePath = @"C:\Images\sample.jpg";

        // Create a new empty document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image from the specified file path (inline, 100% scale)
        builder.InsertImage(imagePath);

        // Save the resulting document
        doc.Save(@"C:\Output\ImageInserted.docx");
    }
}
