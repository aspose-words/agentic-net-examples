using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a custom list based on a predefined template.
        // This adds the list to the document's ListCollection.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list.
        ListLevel level0 = customList.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        level0.Font.Color = Color.DarkBlue;
        // Use Arabic numbering (1, 2, 3, ...) instead of the non‑existent Decimal style.
        level0.NumberStyle = NumberStyle.Arabic;
        level0.StartAt = 1;
        level0.NumberFormat = "%1.";
        level0.NumberPosition = -18;   // Position of the number.
        level0.TextPosition = 36;      // Position of the text after the number.
        level0.TabPosition = 36;
        level0.TrailingCharacter = ListTrailingCharacter.Tab;

        // Configure the second level of the list.
        ListLevel level1 = customList.ListLevels[1];
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.Font.Color = Color.DarkGreen;
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.StartAt = 1;
        level1.NumberFormat = "%2)";
        level1.NumberPosition = 18;
        level1.TextPosition = 72;
        level1.TabPosition = 72;
        level1.TrailingCharacter = ListTrailingCharacter.Space;
        level1.Alignment = ListLevelAlignment.Right;

        // Use DocumentBuilder to add paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the custom list to the builder.
        builder.ListFormat.List = customList;

        // First level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Indent to second level.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");

        // Return to first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("First level item 3");

        // Remove list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file in the current directory.
        doc.Save("CustomList.docx");
    }
}
