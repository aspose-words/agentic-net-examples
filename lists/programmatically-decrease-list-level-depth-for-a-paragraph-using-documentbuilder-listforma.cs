using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListIndentExample
{
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
            builder.Writeln("Item 1");

            // Increase the list level to create a sub‑list.
            builder.ListFormat.ListIndent();
            builder.Writeln("Sub‑item 1");
            builder.Writeln("Sub‑item 2");

            // Decrease the list level only if we are deeper than the top level.
            if (builder.ListFormat.ListLevelNumber > 0)
            {
                // Decrease list level by one.
                builder.ListFormat.ListOutdent();
            }

            // Continue with items at the original level.
            builder.Writeln("Item 2");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to disk.
            doc.Save("ListOutdentExample.docx");
        }
    }
}
