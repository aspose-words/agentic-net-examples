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

        // Create a custom list based on the default numbered template.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level (level 0).
        ListLevel level0 = customList.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 14;
        level0.Font.Color = Color.DarkBlue;
        // Use Arabic numbering (1, 2, 3, ...) instead of the non‑existent Decimal style.
        level0.NumberStyle = NumberStyle.Arabic;
        level0.StartAt = 5;                     // Start numbering at 5.
        level0.NumberFormat = "%1.";            // Custom format.
        level0.NumberPosition = -18;            // Position of the number.
        level0.TextPosition = 36;               // Position of the text.
        level0.TabPosition = 36;

        // Configure the second list level (level 1).
        ListLevel level1 = customList.ListLevels[1];
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;
        level1.Font.Color = Color.DarkGreen;
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.StartAt = 1;
        level1.NumberFormat = "%2)";            // Custom format for second level.
        level1.NumberPosition = 18;
        level1.TextPosition = 72;
        level1.TabPosition = 72;
        level1.TrailingCharacter = ListTrailingCharacter.Space;

        // Use DocumentBuilder to add paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Custom list demonstration:");
        builder.ListFormat.List = customList;   // Apply the custom list.

        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        builder.ListFormat.ListIndent();        // Switch to second level.
        builder.Writeln("Second level subitem 1");
        builder.Writeln("Second level subitem 2");

        builder.ListFormat.ListOutdent();       // Return to first level.
        builder.Writeln("First level item 3");

        builder.ListFormat.RemoveNumbers();     // End the list.

        // Save the document to a file.
        doc.Save("CustomList.docx");
    }
}
