using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1 - level 0");

        // Increase the list level (indent) to create a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 2 - level 1");

        // Increase again to a deeper level.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 3 - level 2");

        // Decrease the list level (outdent) to promote the paragraph.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 4 - back to level 1");

        // Decrease once more to return to the top level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 5 - back to level 0");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("DecreaseIndent.docx");
    }
}
