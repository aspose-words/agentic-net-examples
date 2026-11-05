using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListRestartExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- First section ----------
            // Create a new list for the first section.
            List firstList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Apply the list to the builder.
            builder.ListFormat.List = firstList;
            builder.ListFormat.ListLevelNumber = 1;

            // Write list items.
            builder.Writeln("First section - Item 1");
            builder.Writeln("First section - Item 2");

            // End the first list.
            builder.ListFormat.RemoveNumbers();

            // Add a blank paragraph between sections.
            builder.Writeln();

            // ---------- Second section ----------
            // Create a new list for the second section (new instance restarts numbering).
            List secondList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Apply the second list.
            builder.ListFormat.List = secondList;
            builder.ListFormat.ListLevelNumber = 1;

            // Write list items.
            builder.Writeln("Second section - Item 1");
            builder.Writeln("Second section - Item 2");
            builder.Writeln("Second section - Item 3");

            // End the second list.
            builder.ListFormat.RemoveNumbers();

            // Save the document.
            doc.Save("ListRestartExample.docx");

            // Validation: count paragraphs that are part of a list.
            int listParagraphCount = 0;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");
        }
    }
}
