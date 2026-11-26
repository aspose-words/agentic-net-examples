using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list.
        ListLevel level0 = list.ListLevels[0];
        // Set a custom tab stop (72 points = 1 inch) so text aligns after the number.
        level0.TabPosition = 72;
        // Optional: adjust the position of the number itself.
        level0.NumberPosition = 36;
        level0.Alignment = ListLevelAlignment.Left;

        // Apply the list to the builder.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0; // Use first level.

        // Write list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Validate that three paragraphs are list items.
        int listParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                               .Cast<Paragraph>()
                               .Count(p => p.ListFormat.IsListItem);
        Console.WriteLine($"Number of list paragraphs: {listParagraphs}");

        // Save the document.
        doc.Save("ListWithCustomTab.docx");
    }
}
