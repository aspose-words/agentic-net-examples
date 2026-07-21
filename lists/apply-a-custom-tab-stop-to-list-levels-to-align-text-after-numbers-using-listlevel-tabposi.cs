using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom list based on the default numbered template.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level.
        ListLevel level0 = customList.ListLevels[0];
        level0.Font.Color = System.Drawing.Color.DarkBlue;
        level0.Font.Size = 12;
        level0.NumberStyle = NumberStyle.Arabic;
        level0.StartAt = 1;
        level0.NumberFormat = "%1.";
        // Position of the number (negative moves it left of the left indent).
        level0.NumberPosition = -36;
        // Position where the text starts after the number.
        level0.TextPosition = 144;
        // Use a tab as the separator between number and text.
        level0.TrailingCharacter = ListTrailingCharacter.Tab;
        // Set the tab stop that aligns the text after the number.
        level0.TabPosition = 144;

        // Configure the second list level (optional, demonstrates nesting).
        ListLevel level1 = customList.ListLevels[1];
        level1.Font.Color = System.Drawing.Color.DarkGreen;
        level1.Font.Size = 12;
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.NumberFormat = "%2.";
        level1.NumberPosition = -18;
        level1.TextPosition = 216;
        level1.TrailingCharacter = ListTrailingCharacter.Tab;
        level1.TabPosition = 216;

        // Use DocumentBuilder to add list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = customList;

        // First level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Indent to second level.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item 1");
        builder.Writeln("Second level item 2");
        builder.ListFormat.ListOutdent();

        // Remove list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "CustomListTabPosition.docx");
        doc.Save(outputPath);
    }
}
