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

            // Write the first list item (level 0).
            builder.Writeln("Item at level 0");

            // Increase the list level inside a loop and add items at each deeper level.
            for (int i = 1; i <= 5; i++)
            {
                // Increase the current list level by one.
                builder.ListFormat.ListIndent();

                // Write a paragraph that will appear at the new list level.
                builder.Writeln($"Item at level {i}");
            }

            // Optional: return to the original level before finishing the list.
            for (int i = 0; i < 5; i++)
                builder.ListFormat.ListOutdent();

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the file system.
            doc.Save("IncreaseIndent.docx");
        }
    }
}
