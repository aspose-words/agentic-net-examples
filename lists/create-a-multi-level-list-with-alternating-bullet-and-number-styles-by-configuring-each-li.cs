using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a multilevel list based on the default numbered template.
        // We will reconfigure each level to alternate between numbers and bullets.
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // ---------- Level 0 (top level) – Numbered ----------
        ListLevel level0 = multiLevelList.ListLevels[0];
        level0.NumberStyle = NumberStyle.Arabic;          // Arabic numbers (1, 2, 3, …)
        level0.NumberFormat = "%0.";                     // Default numeric format
        level0.Font.Name = "Times New Roman";
        level0.Font.Size = 12;
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberPosition = -18;                     // Position of the number
        level0.TextPosition = 18;                        // Position of the text after the number
        level0.TrailingCharacter = ListTrailingCharacter.Tab;

        // ---------- Level 1 – Bullet ----------
        ListLevel level1 = multiLevelList.ListLevels[1];
        level1.NumberStyle = NumberStyle.Bullet;         // Bullet style
        level1.NumberFormat = "\u2022";                  // Unicode bullet character (•)
        level1.Font.Name = "Symbol";
        level1.Font.Size = 12;
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberPosition = -18;
        level1.TextPosition = 18;
        level1.TrailingCharacter = ListTrailingCharacter.Tab;

        // ---------- Level 2 – Numbered ----------
        ListLevel level2 = multiLevelList.ListLevels[2];
        level2.NumberStyle = NumberStyle.Arabic;
        level2.NumberFormat = "%0.";                     // Same numeric format as level 0
        level2.Font.Name = "Times New Roman";
        level2.Font.Size = 12;
        level2.Alignment = ListLevelAlignment.Left;
        level2.NumberPosition = -36;
        level2.TextPosition = 36;
        level2.TrailingCharacter = ListTrailingCharacter.Tab;

        // ---------- Level 3 – Bullet ----------
        ListLevel level3 = multiLevelList.ListLevels[3];
        level3.NumberStyle = NumberStyle.Bullet;
        level3.NumberFormat = "\u2022";
        level3.Font.Name = "Symbol";
        level3.Font.Size = 12;
        level3.Alignment = ListLevelAlignment.Left;
        level3.NumberPosition = -36;
        level3.TextPosition = 36;
        level3.TrailingCharacter = ListTrailingCharacter.Tab;

        // Use DocumentBuilder to add paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = multiLevelList;

        // Top‑level items (numbered)
        builder.Writeln("Item 1 – Level 0");
        builder.Writeln("Item 2 – Level 0");

        // Indent to level 1 (bulleted)
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 1 – Level 1");
        builder.Writeln("Item 2 – Level 1");

        // Indent to level 2 (numbered)
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 1 – Level 2");
        builder.Writeln("Item 2 – Level 2");

        // Indent to level 3 (bulleted)
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 1 – Level 3");
        builder.Writeln("Item 2 – Level 3");

        // Outdent back to top level
        builder.ListFormat.ListOutdent(); // back to level 2
        builder.ListFormat.ListOutdent(); // back to level 1
        builder.ListFormat.ListOutdent(); // back to level 0

        // Finish the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "MultiLevelList.docx");
        doc.Save(outputPath);
    }
}
