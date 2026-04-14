using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a new list based on a predefined template and add it to the document's list collection.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level.
        ListLevel level0 = customList.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;
        level0.Font.Color = Color.DarkBlue;
        level0.NumberStyle = NumberStyle.OrdinalText;
        level0.StartAt = 5;
        level0.NumberFormat = "\x0000"; // custom number format
        level0.NumberPosition = -18;
        level0.TextPosition = 36;
        level0.TabPosition = 36;

        // Configure the second list level.
        ListLevel level1 = customList.ListLevels[1];
        level1.Alignment = ListLevelAlignment.Right;
        level1.NumberStyle = NumberStyle.Bullet;
        level1.Font.Name = "Wingdings";
        level1.Font.Color = Color.Green;
        level1.Font.Size = 14;
        level1.NumberFormat = "\xf0af"; // star-shaped bullet
        level1.TrailingCharacter = ListTrailingCharacter.Space;
        level1.NumberPosition = 144;

        // Add paragraphs using the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = customList;
        builder.Writeln("Custom list item 1");
        builder.Writeln("Custom list item 2");
        builder.ListFormat.ListIndent(); // switch to second level
        builder.Writeln("Second level item");
        builder.ListFormat.ListOutdent(); // back to first level
        builder.Writeln("Custom list item 3");
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomList.docx");
        doc.Save(outputPath);
    }
}
