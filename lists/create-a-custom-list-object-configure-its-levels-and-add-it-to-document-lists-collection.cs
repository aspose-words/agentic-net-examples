using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a custom list based on the default numbered template.
        // The Add method automatically adds the list to the document's ListCollection.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list (level 0).
        ListLevel level0 = customList.ListLevels[0];
        level0.Font.Name = "Arial";
        level0.Font.Color = Color.DarkRed;
        level0.Font.Size = 14;
        level0.NumberStyle = NumberStyle.OrdinalText; // e.g., "First", "Second", ...
        level0.StartAt = 1;
        level0.NumberFormat = "\x0000"; // Custom format placeholder.
        level0.NumberPosition = -18;    // Position of the number (negative = left of text).
        level0.TextPosition = 36;       // Position where the text starts.
        level0.TabPosition = 36;        // Tab stop after the number.

        // Configure the second level of the list (level 1) as a bullet.
        ListLevel level1 = customList.ListLevels[1];
        level1.Alignment = ListLevelAlignment.Right;
        level1.NumberStyle = NumberStyle.Bullet;
        level1.Font.Name = "Wingdings";
        level1.Font.Color = Color.Blue;
        level1.Font.Size = 12;
        level1.NumberFormat = "\xf0af"; // Star-shaped bullet.
        level1.TrailingCharacter = ListTrailingCharacter.Space;
        level1.NumberPosition = 144;

        // Use DocumentBuilder to add some paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the custom list to the builder.
        builder.ListFormat.List = customList;

        // First level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Indent to second level.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level bullet 1");
        builder.Writeln("Second level bullet 2");

        // Return to first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("First level item 3");

        // Remove list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomList.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
