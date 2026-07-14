using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel list based on the default numbered template (contains 9 levels).
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure each level individually to alternate between bullet and number styles.
        for (int i = 0; i < 4; i++) // We'll demonstrate the first four levels.
        {
            ListLevel level = multiLevelList.ListLevels[i];

            // Even index (0,2,...) -> bullet style.
            if (i % 2 == 0)
            {
                level.NumberStyle = NumberStyle.Bullet;
                // Use a standard bullet character (•). Unicode 2022.
                level.NumberFormat = "\x2022";
                level.Font.Name = "Symbol";
                level.Font.Color = Color.DarkBlue;
            }
            // Odd index (1,3,...) -> numbered style.
            else
            {
                level.NumberStyle = NumberStyle.Arabic;
                // Use the placeholder for the current level number followed by a period.
                level.NumberFormat = "\x0000.";
                level.Font.Name = "Times New Roman";
                level.Font.Color = Color.DarkRed;
            }

            // Common indentation settings for readability.
            level.NumberPosition = -36;   // Position of the bullet/number.
            level.TextPosition = 144;     // Position where the text starts.
            level.TabPosition = 144;      // Tab stop after the label.
        }

        // Apply the custom list to the builder.
        builder.ListFormat.List = multiLevelList;

        // Add sample items for each configured level.
        for (int i = 0; i < 4; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // Set the current list level.
            builder.Writeln($"Item at level {i + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("MultiLevelAlternatingList.docx");
    }
}
