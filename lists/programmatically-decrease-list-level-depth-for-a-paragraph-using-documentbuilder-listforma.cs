using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // First level items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Increase the list level (indent) to create a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 2.1");
        builder.Writeln("Subitem 2.2");

        // Conditionally decrease the list level (outdent) only if we are not already at the top level.
        if (builder.ListFormat.ListLevelNumber > 0)
        {
            builder.ListFormat.ListOutdent(); // Decrease indent by one level.
        }

        // Continue adding items at the (now) first level.
        builder.Writeln("Item 3");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("Lists_DecreaseIndent.docx");
    }
}
