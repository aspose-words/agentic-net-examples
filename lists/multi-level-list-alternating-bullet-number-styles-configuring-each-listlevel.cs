using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel list based on a default template (contains 9 levels).
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure each level: even indices (0,2,4,…) will be bullets,
        // odd indices (1,3,5,…) will be Arabic numbers.
        for (int i = 0; i < multiLevelList.ListLevels.Count; i++)
        {
            ListLevel level = multiLevelList.ListLevels[i];

            if (i % 2 == 0) // Bullet level.
            {
                level.NumberStyle = NumberStyle.Bullet;
                // Use the standard bullet character (•).
                level.NumberFormat = "\u2022";
                level.Font.Name = "Symbol";
                level.Font.Size = 12;
                level.TrailingCharacter = ListTrailingCharacter.Tab;
            }
            else // Numbered level.
            {
                level.NumberStyle = NumberStyle.Arabic;
                // "\x0000." inserts the current level number followed by a dot.
                level.NumberFormat = "\x0000.";
                level.Font.Name = "Times New Roman";
                level.Font.Size = 12;
                level.TrailingCharacter = ListTrailingCharacter.Tab;
            }

            // Common indentation settings for all levels.
            level.NumberPosition = -18; // Position of the bullet/number.
            level.TextPosition = 36;    // Position where the text starts.
            level.TabPosition = 36;
        }

        // Write sample items for the first six levels to demonstrate the alternating styles.
        builder.Writeln("Multi‑level list with alternating bullet/number styles:");
        for (int level = 0; level < 6; level++)
        {
            builder.ListFormat.List = multiLevelList;
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Level {level + 1} item");
        }

        // End list formatting for subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("MultiLevelAlternatingList.docx");
    }
}
