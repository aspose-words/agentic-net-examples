using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the original list using a built‑in template.
        List originalList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the original list to a few paragraphs.
        builder.ListFormat.List = originalList;
        builder.Writeln("Original Item 1");
        builder.Writeln("Original Item 2");
        builder.Writeln("Original Item 3");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Create a new list that will act as a clone of the original.
        List clonedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Modify the start number of the first level of the cloned list.
        clonedList.ListLevels[0].StartAt = 10;

        // Insert a new section to demonstrate the cloned list.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Apply the cloned list to new paragraphs.
        builder.ListFormat.List = clonedList;
        builder.Writeln("Cloned Item A");
        builder.Writeln("Cloned Item B");
        builder.Writeln("Cloned Item C");
        // End the cloned list.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of any list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");

        // Save the document.
        doc.Save("ClonedListExample.docx");
    }
}
