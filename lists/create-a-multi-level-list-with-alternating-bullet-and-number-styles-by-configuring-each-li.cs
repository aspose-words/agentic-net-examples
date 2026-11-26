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

        // Create a new list based on a bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Level 0 – bullet.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberStyle = NumberStyle.Bullet;
        level0.Font.Name = "Symbol";
        level0.Font.Size = 12;
        level0.NumberFormat = "\u2022";

        // Level 1 – decimal numbering (Arabic).
        ListLevel level1 = list.ListLevels[1];
        level1.NumberStyle = NumberStyle.Arabic;
        level1.NumberFormat = "%1.";
        level1.Font.Size = 12;

        // Level 2 – bullet again.
        ListLevel level2 = list.ListLevels[2];
        level2.NumberStyle = NumberStyle.Bullet;
        level2.Font.Name = "Symbol";
        level2.Font.Size = 12;
        level2.NumberFormat = "\u2022";

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Level 0 item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1");

        // Level 1 item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Subitem 1.1");

        // Level 2 item.
        builder.ListFormat.ListLevelNumber = 2;
        builder.Writeln("Subsubitem 1.1.1");

        // Back to Level 0.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 2");

        // Level 1 under Item 2.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Subitem 2.1");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are list items.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"List paragraphs count: {listParagraphCount}");

        // Save the document.
        doc.Save("MultiLevelList.docx");
    }
}
