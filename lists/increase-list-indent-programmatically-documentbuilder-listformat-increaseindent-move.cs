using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item at level 0");

        // Increase the list level – this is the core operation requested.
        builder.ListFormat.ListIndent(); // level becomes 1
        builder.Writeln("Item at level 1");

        // Increase again to demonstrate deeper nesting.
        builder.ListFormat.ListIndent(); // level becomes 2
        builder.Writeln("Item at level 2");

        // Optionally outdent back to a previous level.
        builder.ListFormat.ListOutdent(); // level back to 1
        builder.Writeln("Back to level 1");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ListIndentExample.docx");
    }
}
