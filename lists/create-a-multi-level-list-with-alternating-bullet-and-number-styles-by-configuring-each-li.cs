using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class MultiLevelListExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new multilevel list based on a default numbered template.
        // All lists created this way contain 9 levels.
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure each level: even levels use bullets, odd levels use Arabic numbers.
        for (int i = 0; i < multiLevelList.ListLevels.Count; i++)
        {
            ListLevel level = multiLevelList.ListLevels[i];

            // Alternate between bullet and number styles.
            if (i % 2 == 0) // Even level – bullet.
            {
                level.NumberStyle = NumberStyle.Bullet;
                // Use a simple bullet character. The NumberFormat can be any character; here we use the standard bullet.
                level.NumberFormat = "\u2022";
                level.Font.Name = "Symbol";
                level.Font.Size = 12;
                level.TrailingCharacter = ListTrailingCharacter.Space;
            }
            else // Odd level – Arabic number.
            {
                level.NumberStyle = NumberStyle.Arabic;
                level.NumberFormat = "\x0000."; // Placeholder for the current level number.
                level.Font.Name = "Times New Roman";
                level.Font.Size = 12;
                level.TrailingCharacter = ListTrailingCharacter.Tab;
            }

            // Optional: set indentation so that each level is indented further.
            level.NumberPosition = -18; // Position of the bullet/number.
            level.TextPosition = 36;    // Position where the text starts.
            level.TabPosition = 36;     // Tab stop after the label.
        }

        // Apply the list to the builder.
        builder.ListFormat.List = multiLevelList;

        // Add sample items for three levels, two items per level.
        for (int level = 0; level < 3; level++)
        {
            // Set the current list level.
            builder.ListFormat.ListLevelNumber = level;

            // Write two items at this level.
            builder.Writeln($"Item {level + 1}.1");
            builder.Writeln($"Item {level + 1}.2");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output folder.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MultiLevelList.docx");
        doc.Save(outputPath);
    }
}
