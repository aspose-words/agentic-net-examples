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

            // Create a numbered list using a built‑in list template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Apply the list to the builder.
            builder.ListFormat.List = list;

            // First level items (level 0).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // Sub‑level items (level 1).
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("Subitem 2.1");
            builder.Writeln("Subitem 2.2");

            // Return to the original level (level 0).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("Item 3");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Validate the number of list paragraphs.
            int listParagraphCount = 0;
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");

            // Save the document to a file.
            doc.Save("ListIndent.docx");
        }
    }
}
