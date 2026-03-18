using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class NineLevelListExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list based on a built‑in template.
        // All lists created with Aspose.Words contain 9 levels.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define formatting for each of the nine levels (0‑8).
        for (int i = 0; i < 9; i++)
        {
            ListLevel level = list.ListLevels[i];

            // Example: use a different font size for each level.
            level.Font.Name = "Arial";
            level.Font.Size = 12 + i;               // 12pt, 13pt, … 20pt
            level.Font.Color = Color.Black;
            level.Font.Bold = i % 2 == 0;            // Bold on even levels

            // Alternate number styles: Arabic, LowercaseLetter, LowercaseRoman, etc.
            switch (i % 4)
            {
                case 0: level.NumberStyle = NumberStyle.Arabic; break;
                case 1: level.NumberStyle = NumberStyle.LowercaseLetter; break;
                case 2: level.NumberStyle = NumberStyle.LowercaseRoman; break;
                case 3: level.NumberStyle = NumberStyle.Bullet; break;
            }

            // Use the placeholder \x0000 for the current level number.
            // For bullet levels we supply a custom symbol.
            if (level.NumberStyle == NumberStyle.Bullet)
                level.NumberFormat = "\u2022"; // standard bullet
            else
                level.NumberFormat = "\x0000.";

            // Position the number/bullet and the text.
            level.NumberPosition = -36 * (i + 1);   // shift left for deeper levels
            level.TextPosition = 36 * (i + 1);      // indent text
            level.TabPosition = level.TextPosition; // align tab stop

            // Use a space after the number/bullet.
            level.TrailingCharacter = ListTrailingCharacter.Space;

            // Align numbers to the left.
            level.Alignment = ListLevelAlignment.Left;
        }

        // Demonstrate the nine‑level hierarchy.
        builder.Writeln("Nine‑level hierarchical list:");
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.List = list;          // Apply the list to the paragraph.
            builder.ListFormat.ListLevelNumber = i;  // Select the current level (0‑8).
            builder.Writeln($"Level {i + 1} item");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("NineLevelList.docx");
    }
}
