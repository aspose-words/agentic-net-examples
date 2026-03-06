using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document (uses the Document constructor rule)
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document (allows content insertion)
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a placeholder paragraph where a group shape can be inserted later
        builder.Writeln("Placeholder for group shape insertion:");

        // Save the document as a DOCX file (uses the Document.Save rule)
        doc.Save("GroupShapeReady.docx");
    }
}
