using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace HierarchicalListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a multilevel list based on the default numbered template.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);

            // Define custom properties for each of the nine list levels (0‑8).
            for (int i = 0; i < 9; i++)
            {
                ListLevel level = list.ListLevels[i];

                // Example customizations – you can adjust as needed.
                level.Font.Name = "Times New Roman";
                level.Font.Size = 12 + i;                     // Increment size per level.
                level.Font.Color = Color.FromArgb(20 * i, 0, 255 - 20 * i);
                level.Alignment = ListLevelAlignment.Left;
                level.NumberStyle = NumberStyle.Arabic;       // Use Arabic numbers for simplicity.
                level.NumberFormat = "\\x0000.";              // Display the number followed by a dot.
                level.NumberPosition = -36 - (i * 10);        // Shift number position left for deeper levels.
                level.TextPosition = 144 + (i * 10);          // Indent text for each level.
                level.TabPosition = 144 + (i * 10);
                level.TrailingCharacter = ListTrailingCharacter.Tab;
            }

            // Write sample paragraphs for each level to demonstrate the hierarchy.
            builder.ListFormat.List = list;
            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i; // Set current list level.
                builder.Writeln($"Level {i + 1} item");
            }

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the document to the file system.
            doc.Save("HierarchicalList.docx");
        }
    }
}
