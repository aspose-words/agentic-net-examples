using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace ListIndentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list using a ListTemplate.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // Write list items, increasing the indent (list level) in each loop iteration.
            // ListLevelNumber is zero‑based: 0 = top level, 1 = first sub‑level, etc.
            for (int i = 1; i <= 5; i++)
            {
                // Set the desired list level before writing the paragraph.
                builder.ListFormat.ListLevelNumber = i - 1;

                // Write the current list item.
                builder.Writeln($"Item {i} at level {builder.ListFormat.ListLevelNumber}");
            }

            // Remove list formatting after the list is complete.
            builder.ListFormat.RemoveNumbers();

            // Validation: count the number of paragraphs that are part of the list.
            int listParagraphCount = 0;
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            Console.WriteLine($"Number of list paragraphs: {listParagraphCount}");

            // Save the document to a file.
            doc.Save("ListIndentExample.docx");
        }
    }
}
