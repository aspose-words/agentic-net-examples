using System;
using System.Drawing;
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
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define custom properties for each of the nine list levels (0‑8).
        for (int i = 0; i < list.ListLevels.Count; i++)
        {
            ListLevel level = list.ListLevels[i];

            // Example font customizations.
            level.Font.Name = "Arial";
            level.Font.Size = 12 + i; // increase size with depth.
            level.Font.Color = Color.FromArgb(20 * i, 0, 255 - 20 * i);

            // Use different number styles for illustration.
            switch (i % 4)
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
                    level.NumberStyle = NumberStyle.Bullet;
                    level.NumberFormat = "\u2022"; // bullet character.
                    break;
            }

            // Adjust indent positions for each level.
            level.NumberPosition = -36 - i * 10;
            level.TextPosition = 144 + i * 20;
            level.TabPosition = 144 + i * 20;
        }

        // Apply the list to the builder and write a paragraph for each level.
        builder.ListFormat.List = list;
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i;
            builder.Writeln($"Level {i + 1}");
        }

        // End list formatting.
        builder.ListFormat.List = null;

        // Save the document to the current directory.
        doc.Save("HierarchicalList.docx");
    }
}
