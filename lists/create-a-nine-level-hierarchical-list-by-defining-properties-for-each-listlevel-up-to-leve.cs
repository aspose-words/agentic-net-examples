using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListHierarchyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a multilevel list based on the default numbered template.
            // All Aspose.Words lists contain 9 levels (0‑8).
            List multilevelList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Define custom formatting for each of the nine list levels.
            for (int level = 0; level < 9; level++)
            {
                ListLevel listLevel = multilevelList.ListLevels[level];

                // Example customizations:
                listLevel.Font.Name = "Arial";
                listLevel.Font.Size = 12 + level;               // Increment size per level.
                listLevel.Font.Color = Color.FromArgb(20 * level, 0, 255 - 20 * level);
                listLevel.NumberStyle = NumberStyle.Arabic;    // Use Arabic numbers for all levels.
                listLevel.NumberFormat = $"%{level}.";         // Simple format showing the level index.
                listLevel.NumberPosition = -36 - (level * 10); // Indent numbers further for deeper levels.
                listLevel.TextPosition = 144 + (level * 20);   // Indent text after the number.
                listLevel.TabPosition = listLevel.TextPosition;
                listLevel.Alignment = ListLevelAlignment.Left;
                listLevel.TrailingCharacter = ListTrailingCharacter.Tab;
            }

            // Apply the list to the builder and write one item per level.
            builder.ListFormat.List = multilevelList;

            for (int level = 0; level < 9; level++)
            {
                // Set the current list level (0‑8) for the paragraph.
                builder.ListFormat.ListLevelNumber = level;

                // Write a paragraph that will appear at this level.
                builder.Writeln($"Level {level + 1}");
            }

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the local file system.
            string outputPath = "NineLevelHierarchy.docx";
            doc.Save(outputPath);
        }
    }
}
