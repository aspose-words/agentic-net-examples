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

        // Create a new list that will hold nine hierarchical levels.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define properties for each of the nine levels (0‑based index).
        for (int i = 0; i < 9; i++)
        {
            ListLevel level = list.ListLevels[i];

            // Use Arabic numerals for all levels.
            level.NumberStyle = NumberStyle.Arabic;

            // Define a simple number format, e.g., "1.", "2.", etc.
            level.NumberFormat = $"{i + 1}.";

            // Align numbers to the left.
            level.Alignment = ListLevelAlignment.Left;

            // Set the font for the list numbers.
            level.Font.Name = "Arial";
            level.Font.Size = 12;

            // Start numbering at 1 for each level.
            level.StartAt = 1;
        }

        // Add a paragraph for each level, applying the corresponding list level.
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = i; // Set current level (0‑based).
            builder.Writeln($"Level {i + 1} item");
            builder.ListFormat.RemoveNumbers(); // Reset list formatting for next paragraph.
        }

        // Save the document to disk.
        doc.Save("HierarchicalList.docx");
    }
}
