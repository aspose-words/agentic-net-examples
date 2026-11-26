using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level to use Upper Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;
        level.NumberFormat = "%1.";
        level.Alignment = ListLevelAlignment.Left;

        // Apply the list to the builder.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0;

        // Write list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of the list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"List paragraphs count: {listParagraphCount}");

        // Save the document to a file.
        doc.Save("ListUpperRoman.docx");
    }
}
