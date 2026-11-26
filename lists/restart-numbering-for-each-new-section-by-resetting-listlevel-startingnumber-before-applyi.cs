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

        // ---------- First section ----------
        // Create a list and set its starting number.
        List firstList = doc.Lists.Add(ListTemplate.NumberDefault);
        // The correct property to set the starting number is StartAt.
        firstList.ListLevels[0].StartAt = 1;

        // Apply the list to the builder.
        builder.ListFormat.List = firstList;
        builder.ListFormat.ListLevelNumber = 0;

        // Write list items.
        builder.Writeln("Item 1.1");
        builder.Writeln("Item 1.2");
        builder.Writeln("Item 1.3");

        // End the first list.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph between sections.
        builder.Writeln("Some text between sections.");

        // ---------- Second section ----------
        // Create a new list instance to restart numbering.
        List secondList = doc.Lists.Add(ListTemplate.NumberDefault);
        secondList.ListLevels[0].StartAt = 1; // restart at 1

        // Apply the new list.
        builder.ListFormat.List = secondList;
        builder.ListFormat.ListLevelNumber = 0;

        // Write list items.
        builder.Writeln("Item 2.1");
        builder.Writeln("Item 2.2");
        builder.Writeln("Item 2.3");

        // End the second list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("RestartNumbering.docx");

        // Simple validation: count paragraphs that are part of a list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"Document created with {listParagraphCount} list paragraphs.");
    }
}
