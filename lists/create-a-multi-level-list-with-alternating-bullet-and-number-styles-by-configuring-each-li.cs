using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a multilevel list based on a default numbered template.
        // All 9 levels are created automatically.
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure each level: even levels -> bullet, odd levels -> Arabic number.
        for (int i = 0; i < multiLevelList.ListLevels.Count; i++)
        {
            ListLevel level = multiLevelList.ListLevels[i];

            if (i % 2 == 0) // Bullet level
            {
                level.NumberStyle = NumberStyle.Bullet;
                level.Font.Name = "Wingdings";
                // Unicode bullet character (•) – you can use any symbol you prefer.
                level.NumberFormat = "\x2022";
                level.TrailingCharacter = ListTrailingCharacter.Space;
            }
            else // Numbered level
            {
                level.NumberStyle = NumberStyle.Arabic;
                // Placeholder for the current level's number followed by a period.
                level.NumberFormat = "\x0000.";
                level.TrailingCharacter = ListTrailingCharacter.Space;
            }

            // Optional: set indent positions so the list looks tidy.
            level.NumberPosition = -18;   // Position of the bullet/number.
            level.TextPosition = 18;      // Position of the text after the bullet/number.
            level.TabPosition = 36;       // Tab stop for wrapped lines.
        }

        // Use DocumentBuilder to add paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = multiLevelList;

        // Add items for the first four levels to demonstrate alternating styles.
        for (int level = 0; level < 4; level++)
        {
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Item at level {level}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("MultiLevelList.docx");
    }
}
