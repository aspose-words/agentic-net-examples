using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel list based on the default numbered template.
        // All lists created this way contain 9 levels.
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure each level: even levels will use bullets, odd levels will use numbers.
        for (int i = 0; i < multiLevelList.ListLevels.Count; i++)
        {
            ListLevel level = multiLevelList.ListLevels[i];

            if (i % 2 == 0) // Bullet level
            {
                level.NumberStyle = NumberStyle.Bullet;
                // Use a standard bullet character. You can also use a Wingdings character if desired.
                level.NumberFormat = "\u2022"; // •
                level.Font.Name = "Symbol";
            }
            else // Numbered level
            {
                level.NumberStyle = NumberStyle.Arabic;
                // Use the placeholder for the current level number.
                level.NumberFormat = "\x0000";
                level.Font.Name = "Times New Roman";
            }

            // Optional: adjust indent positions for better readability.
            level.NumberPosition = -18; // Position of the bullet/number.
            level.TextPosition = 18;    // Position where the text starts.
            level.TabPosition = 36;     // Tab stop after the label.
        }

        // Apply the list to the builder and write sample items for each level.
        builder.ListFormat.List = multiLevelList;

        for (int level = 0; level < multiLevelList.ListLevels.Count; level++)
        {
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Item at level {level + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("MultiLevelList.docx");
    }
}
