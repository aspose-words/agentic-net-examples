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
        builder.Writeln("Item 1");

        // Increase the list level to create a sub‑item.
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 1.1");

        // Decrease the list level only if we are not already at the top level.
        if (builder.ListFormat.ListLevelNumber > 0)
        {
            builder.ListFormat.ListOutdent(); // Decreases list depth by one level.
        }

        // Continue the list after outdenting.
        builder.Writeln("Item 2 after outdent");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("DecreaseListLevel.docx");
    }
}
