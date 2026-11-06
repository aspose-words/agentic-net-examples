using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define properties for each of the nine levels.
        // Only NumberFormat and Font size are set; the default NumberStyle is used.
        // Level 1
        ListLevel level1 = list.ListLevels[0];
        level1.NumberFormat = "%1.";
        level1.Font.Size = 12;

        // Level 2
        ListLevel level2 = list.ListLevels[1];
        level2.NumberFormat = "%2.";
        level2.Font.Size = 12;

        // Level 3
        ListLevel level3 = list.ListLevels[2];
        level3.NumberFormat = "%3)";
        level3.Font.Size = 12;

        // Level 4
        ListLevel level4 = list.ListLevels[3];
        level4.NumberFormat = "%4)";
        level4.Font.Size = 12;

        // Level 5
        ListLevel level5 = list.ListLevels[4];
        level5.NumberFormat = "%5.";
        level5.Font.Size = 12;

        // Level 6
        ListLevel level6 = list.ListLevels[5];
        level6.NumberFormat = "%6.";
        level6.Font.Size = 12;

        // Level 7
        ListLevel level7 = list.ListLevels[6];
        level7.NumberFormat = "%7)";
        level7.Font.Size = 12;

        // Level 8
        ListLevel level8 = list.ListLevels[7];
        level8.NumberFormat = "%8)";
        level8.Font.Size = 12;

        // Level 9
        ListLevel level9 = list.ListLevels[8];
        level9.NumberFormat = "%9.";
        level9.Font.Size = 12;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write one item for each level, setting the appropriate level number (1‑based).
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i + 1; // Levels are 1‑based.
            builder.Writeln($"Level {i + 1} item");
        }

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("NineLevelList.docx");
    }
}
