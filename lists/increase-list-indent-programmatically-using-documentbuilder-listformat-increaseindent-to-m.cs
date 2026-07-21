using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // First level items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Increase the list indent – the next items will be at a deeper level.
        builder.ListFormat.ListIndent();

        // Second level items.
        builder.Writeln("Subitem 2.1");
        builder.Writeln("Subitem 2.2");

        // Decrease the indent back to the first level.
        builder.ListFormat.ListOutdent();

        // Continue first level items.
        builder.Writeln("Item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        doc.Save("ListIndentExample.docx");
    }
}
