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
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Increase the list level – creates a sub‑list under "Item 2".
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 2.1");
        builder.Writeln("Subitem 2.2");

        // Decrease the list level – promotes the next paragraph back to the parent level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 3");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ListOutdentExample.docx");
    }
}
