using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a numbered list.
            builder.ListFormat.ApplyNumberDefault();

            // Add a base list item.
            builder.Writeln("Base item");

            // Increase list level depth in a loop.
            for (int i = 0; i < 3; i++)
            {
                // Increase the list level (indent) by setting ListLevelNumber.
                builder.ListFormat.ListLevelNumber = builder.ListFormat.ListLevelNumber + 1;

                // Add a paragraph at the new level.
                builder.Writeln($"Sub item level {i + 1}");
            }

            // Reset list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document.
            doc.Save("ListIndentExample.docx");
        }
    }
}
