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

        // First list item at level 0.
        builder.Writeln("Item 1 - level 0");

        // Increase the list level (indent) for the next item.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 2 - level 1");

        // Decrease the list level only if we are deeper than the top level.
        if (builder.ListFormat.ListLevelNumber > 0)
        {
            builder.ListFormat.ListOutdent(); // Decrease indent.
        }

        // Add another item after the outdent.
        builder.Writeln("Item 3 - back to level 0");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("DecreaseListLevel.docx");
    }
}
