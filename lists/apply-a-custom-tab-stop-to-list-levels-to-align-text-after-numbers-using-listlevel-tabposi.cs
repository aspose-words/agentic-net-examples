using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a list based on the built‑in numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level.
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Color = System.Drawing.Color.Red;
        level0.Font.Size = 24;
        level0.NumberStyle = NumberStyle.OrdinalText;
        level0.StartAt = 1;
        level0.NumberFormat = "\x0000";

        // The tab character will be placed after the number.
        level0.TrailingCharacter = ListTrailingCharacter.Tab;

        // Position the number (negative moves it left of the left indent).
        level0.NumberPosition = -36;
        // Position the text after the tab stop.
        level0.TextPosition = 144;
        // Set a custom tab stop that aligns the text after the number.
        level0.TabPosition = 144;

        // Configure the second list level (optional, shows inheritance).
        ListLevel level1 = list.ListLevels[1];
        level1.Alignment = ListLevelAlignment.Right;
        level1.NumberStyle = NumberStyle.Bullet;
        level1.Font.Name = "Wingdings";
        level1.Font.Color = System.Drawing.Color.Blue;
        level1.Font.Size = 24;
        level1.NumberFormat = "\xf0af";
        level1.TrailingCharacter = ListTrailingCharacter.Space;
        level1.NumberPosition = 144;

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // First‑level items.
        builder.Writeln("First item – aligned after the number.");
        builder.Writeln("Second item – aligned after the number.");

        // Indent to second level.
        builder.ListFormat.ListIndent();
        builder.Writeln("Sub‑item at level 2.");
        builder.Writeln("Another sub‑item at level 2.");

        // Return to first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Third item – back to first level.");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("CustomListTabPosition.docx");
    }
}
