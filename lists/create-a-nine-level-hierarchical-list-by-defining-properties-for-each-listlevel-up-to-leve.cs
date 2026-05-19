using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define properties for each of the nine list levels (0‑8).
        for (int i = 0; i < list.ListLevels.Count; i++)
        {
            ListLevel level = list.ListLevels[i];

            // Example formatting: font name, size, and color.
            level.Font.Name = "Arial";
            level.Font.Size = 12 + i; // increase size with level depth
            level.Font.Color = Color.FromArgb(20 * i, 0, 255 - 20 * i);

            // Assign different number styles and formats per level.
            switch (i)
            {
                case 0:
                    level.NumberStyle = NumberStyle.Arabic;
                    level.NumberFormat = "%1.";
                    break;
                case 1:
                    level.NumberStyle = NumberStyle.LowercaseLetter;
                    level.NumberFormat = "%2.";
                    break;
                case 2:
                    level.NumberStyle = NumberStyle.LowercaseRoman;
                    level.NumberFormat = "%3.";
                    break;
                case 3:
                    level.NumberStyle = NumberStyle.UppercaseLetter;
                    level.NumberFormat = "%4.";
                    break;
                case 4:
                    level.NumberStyle = NumberStyle.UppercaseRoman;
                    level.NumberFormat = "%5.";
                    break;
                case 5:
                    level.NumberStyle = NumberStyle.Ordinal;
                    level.NumberFormat = "%6.";
                    break;
                case 6:
                    level.NumberStyle = NumberStyle.OrdinalText;
                    level.NumberFormat = "Level %7";
                    break;
                case 7:
                    level.NumberStyle = NumberStyle.Bullet;
                    level.NumberFormat = "\u2022"; // solid bullet
                    break;
                case 8:
                    level.NumberStyle = NumberStyle.Bullet;
                    level.NumberFormat = "\u25E6"; // white bullet
                    break;
            }

            // Set indent positions for visual hierarchy.
            level.NumberPosition = -36 - i * 10;
            level.TextPosition = 144 + i * 10;
            level.TabPosition = 144 + i * 10;
        }

        // Add a sample paragraph for each level to demonstrate the hierarchy.
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = i;
            builder.Writeln($"Level {i + 1} item");
        }

        // End list formatting.
        builder.ListFormat.List = null;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HierarchicalList.docx");
        doc.Save(outputPath);
    }
}
