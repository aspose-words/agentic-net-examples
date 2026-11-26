using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list and assign it to the builder.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // Write the first (top‑level) list item (level 0).
        builder.Writeln("Item 1");

        // Increase the list level inside a loop to create nested items.
        for (int i = 2; i <= 5; i++)
        {
            // Set the desired list level (zero‑based). Level 1 for Item 2, level 2 for Item 3, etc.
            builder.ListFormat.ListLevelNumber = i - 1;

            // Write a nested list item.
            builder.Writeln($"Item {i}");
        }

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NestedList.docx");
        doc.Save(outputPath);

        // Validation: count paragraphs that are part of the list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.IsListItem)
                listParagraphCount++;
        }

        // Output validation result.
        Console.WriteLine($"Document created with {listParagraphCount} list paragraphs.");
    }
}
