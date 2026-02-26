using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will serve as the insertion point for a future group shape.
        builder.Writeln("Insert group shape here.");

        // Save the document in DOCX format.
        doc.Save("GroupShapeReady.docx");
    }
}
