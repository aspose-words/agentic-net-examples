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

        // Add a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define custom properties for each of the nine list levels (0‑8).
        for (int i = 0; i < list.ListLevels.Count; i++)
        {
            ListLevel level = list.ListLevels[i];
            level.Font.Name = "Arial";
            level.Font.Size = 12 + i; // Increment font size per level.
            level.Font.Color = Color.FromArgb(20 * i, 0, 255 - 20 * i);
            level.NumberStyle = NumberStyle.Arabic;
            level.NumberFormat = $"{i + 1}.";
            level.NumberPosition = -36 - (i * 10);
            level.TextPosition = 144 + (i * 20);
            level.TabPosition = 144 + (i * 20);
        }

        // Apply the list to the builder and write one paragraph for each level.
        builder.ListFormat.List = list;
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i;
            builder.Writeln($"Level {i + 1}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("HierarchicalList.docx");
    }
}
