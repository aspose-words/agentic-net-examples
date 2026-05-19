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

        // Add a couple of top‑level list items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Increase the list level to create a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 1");
        builder.Writeln("Subitem 2");

        // If we are currently indented (level > 0), decrease the list level.
        if (builder.ListFormat.ListLevelNumber > 0)
        {
            // Decrease the list level depth.
            builder.ListFormat.ListOutdent();
        }

        // Continue adding items at the (now) current level.
        builder.Writeln("Item 3");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the file system.
        doc.Save("DecreaseListLevel.docx");
    }
}
