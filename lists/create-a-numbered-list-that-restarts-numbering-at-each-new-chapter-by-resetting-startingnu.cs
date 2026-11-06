using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Chapter 1 ----------
        builder.Writeln("Chapter 1");

        // Create a new list for Chapter 1. A fresh list automatically starts numbering at 1.
        List chapter1List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder and set the level (top level = 0).
        builder.ListFormat.List = chapter1List;
        builder.ListFormat.ListLevelNumber = 0;

        // Write list items for Chapter 1.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the list for Chapter 1.
        builder.ListFormat.RemoveNumbers();

        // ---------- Chapter 2 ----------
        builder.Writeln("Chapter 2");

        // Create a new list for Chapter 2 to restart numbering.
        List chapter2List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the new list.
        builder.ListFormat.List = chapter2List;
        builder.ListFormat.ListLevelNumber = 0;

        // Write list items for Chapter 2.
        builder.Writeln("Item A");
        builder.Writeln("Item B");
        builder.Writeln("Item C");

        // End the list for Chapter 2.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("NumberedListsByChapter.docx");

        // Validation: count paragraphs that are part of a list.
        int listParagraphCount = doc.GetChildNodes(NodeType.Paragraph, true)
                                    .Cast<Paragraph>()
                                    .Count(p => p.ListFormat.IsListItem);

        // Expected number of list items: 6.
        Console.WriteLine($"List paragraphs count: {listParagraphCount}");
    }
}
