using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Create a custom list based on a predefined template.
            // Here we start from the default numbered list which has 9 levels.
            List customList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Configure the first list level (level 0).
            ListLevel level0 = customList.ListLevels[0];
            level0.Font.Color = Color.Red;
            level0.Font.Size = 24;
            level0.NumberStyle = NumberStyle.OrdinalText; // e.g., "First", "Second", ...
            level0.StartAt = 1;
            level0.NumberFormat = "\x0000"; // Custom format placeholder.
            level0.NumberPosition = -36;    // Position of the number relative to the left margin.
            level0.TextPosition = 144;      // Position of the text after the number.
            level0.TabPosition = 144;       // Tab stop after the number.

            // Configure the second list level (level 1) as a bullet list.
            ListLevel level1 = customList.ListLevels[1];
            level1.Alignment = ListLevelAlignment.Right;
            level1.NumberStyle = NumberStyle.Bullet;
            level1.Font.Name = "Wingdings";
            level1.Font.Color = Color.Blue;
            level1.Font.Size = 24;
            level1.NumberFormat = "\xf0af"; // Star-shaped bullet.
            level1.TrailingCharacter = ListTrailingCharacter.Space;
            level1.NumberPosition = 144;

            // Use DocumentBuilder to add paragraphs that use the custom list.
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

            // Outdent back to first level.
            builder.ListFormat.ListOutdent();
            builder.Writeln("First level item 3");

            // Remove list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the local file system.
            doc.Save("CustomList.docx");
        }
    }
}
