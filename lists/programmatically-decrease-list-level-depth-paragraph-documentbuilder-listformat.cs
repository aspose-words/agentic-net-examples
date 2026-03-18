using System;
using Aspose.Words;
using Aspose.Words.Lists;

class DecreaseListLevelExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");

        // Increase the list level – creates a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Sub‑item 1");

        // Conditional: only outdent if we are currently inside a list and the level is greater than zero.
        if (builder.ListFormat.IsListItem && builder.ListFormat.ListLevelNumber > 0)
        {
            // Decrease the list level depth for the next paragraph.
            builder.ListFormat.ListOutdent();
        }

        // This paragraph will be at the previous (higher) list level.
        builder.Writeln("Back to parent level");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("DecreasedListLevel.docx");
    }
}
