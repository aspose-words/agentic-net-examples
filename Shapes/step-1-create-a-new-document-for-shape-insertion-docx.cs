using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class ShapeInsertionDemo
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // (Optional) Insert a simple shape to verify that the document is ready for shape operations.
        // Here we insert an inline rectangle shape of size 100x50 points.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Save the document in DOCX format.
        doc.Save("ShapeInsertion.docx");
    }
}
