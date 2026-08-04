using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Define custom formatting for each of the nine list levels (0‑8).
        for (int i = 0; i < 9; i++)
        {
            ListLevel level = list.ListLevels[i];

            // Example customizations: font, size, color, numbering style and positions.
            level.Font.Name = "Arial";
            level.Font.Size = 12 + i; // Increment size per level.
            level.Font.Color = Color.FromArgb((i * 30) % 256, (i * 60) % 256, (i * 90) % 256);
            level.NumberStyle = NumberStyle.Arabic; // Simple Arabic numbers for all levels.
            level.NumberFormat = "\x0000"; // Use the default number placeholder.
            level.NumberPosition = -36 - i * 5; // Shift number position leftwards.
            level.TextPosition = 144 + i * 10; // Indent text for each level.
            level.TabPosition = 144 + i * 10; // Align tab position with text.
        }

        // Apply the list to the document and write a paragraph for each level.
        builder.ListFormat.List = list;
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // Set current list level.
            builder.Writeln($"Level {i + 1}");
        }

        // Remove list formatting from subsequent paragraphs (optional).
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file in the current directory.
        doc.Save("HierarchicalList.docx");
    }
}
