using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list using a built‑in template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // First level list item.
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("Item 1");

            // Second level (nested) list item.
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("Item 1.1");

            // Conditional decrease of indent (back to previous level).
            bool shouldDecrease = true; // Replace with real condition as needed.
            if (shouldDecrease)
            {
                // Decrease the list level safely.
                if (builder.ListFormat.ListLevelNumber > 0)
                {
                    builder.ListFormat.ListLevelNumber--;
                }
            }

            // This item will be at the level after the possible decrease.
            // Ensure the correct level is set before writing.
            builder.Writeln("Item 2");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document.
            doc.Save("ListIndentExample.docx");

            // Validation: count paragraphs that are part of a list.
            int listParagraphCount = doc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .Count(p => p.ListFormat.IsListItem);

            Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");
        }
    }
}
