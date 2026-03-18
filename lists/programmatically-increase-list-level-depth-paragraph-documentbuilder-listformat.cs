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

        // Write the first (top‑level) list item.
        builder.Writeln("Level 0");

        // Increase the list level inside a loop and add items at each deeper level.
        for (int i = 1; i <= 5; i++)
        {
            // Increase the current list level by one.
            builder.ListFormat.ListIndent();

            // Write an item at the new level.
            builder.Writeln($"Level {i}");
        }

        // Return to the original level (optional) and end the list.
        while (builder.ListFormat.ListLevelNumber > 0)
            builder.ListFormat.ListOutdent();

        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ListIndentExample.docx");
    }
}
