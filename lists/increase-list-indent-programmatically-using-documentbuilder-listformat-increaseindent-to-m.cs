using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1 (level 0)");
        builder.Writeln("Item 2 (level 0)");

        // Increase the list indent – the next paragraph will be at level 1.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 3 (level 1)");
        builder.Writeln("Item 4 (level 1)");

        // Increase the indent again – now at level 2.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 5 (level 2)");

        // Decrease the indent back to level 1.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 6 (back to level 1)");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the file system.
        doc.Save("IncreaseIndent.docx");
    }
}
