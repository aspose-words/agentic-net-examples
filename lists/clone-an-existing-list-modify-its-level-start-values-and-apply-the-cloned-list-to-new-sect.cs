using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Create the original list using a built‑in template.
            // -------------------------------------------------
            List originalList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Apply the original list to paragraphs.
            builder.ListFormat.List = originalList;
            builder.Writeln("Original Item 1");
            builder.Writeln("Original Item 2");
            // End the list.
            builder.ListFormat.RemoveNumbers();

            // -------------------------------------------------
            // Create a new list that will act as the cloned list.
            // -------------------------------------------------
            List clonedList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Modify the start number of the first level in the cloned list.
            // For example, start numbering at 5 instead of 1.
            clonedList.ListLevels[0].StartAt = 5;

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // -------------------------------------------------
            // Apply the cloned list to new paragraphs in the new section.
            // -------------------------------------------------
            builder.ListFormat.List = clonedList;
            builder.Writeln("Cloned Item 1");
            builder.Writeln("Cloned Item 2");
            // End the cloned list.
            builder.ListFormat.RemoveNumbers();

            // -------------------------------------------------
            // Validation: count paragraphs that are part of a list.
            // -------------------------------------------------
            int listParagraphCount = 0;
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            Console.WriteLine($"Total list paragraphs in document: {listParagraphCount}");

            // Save the document.
            doc.Save("ListCloneOutput.docx");
        }
    }
}
