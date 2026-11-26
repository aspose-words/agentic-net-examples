using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace NineLevelListExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a custom list based on a predefined template (NumberDefault) that we will customize.
            List customList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Define properties for each of the nine list levels (1‑9).
            for (int i = 0; i < 9; i++)
            {
                ListLevel level = customList.ListLevels[i];
                // Use Arabic numbering (1., 2., 3., …) for all levels.
                level.NumberStyle = NumberStyle.Arabic;
                // Show the level number followed by a period.
                level.NumberFormat = $"{i + 1}.";
                // Align numbers to the left.
                level.Alignment = ListLevelAlignment.Left;
                // Indent each level further (18 points per level).
                level.NumberPosition = 18 * (i + 1);
                // Set the text indent so the list text aligns after the number.
                level.TextPosition = level.NumberPosition + 18;
                // Use a standard font size for the list numbers.
                level.Font.Size = 12;
            }

            // Write one item for each level, nesting them appropriately.
            for (int levelNumber = 1; levelNumber <= 9; levelNumber++)
            {
                // Assign the custom list to the builder.
                builder.ListFormat.List = customList;
                // Set the current nesting level (1‑based).
                builder.ListFormat.ListLevelNumber = levelNumber;
                // Write the list item text.
                builder.Writeln($"Level {levelNumber} item");
            }

            // Remove list formatting after the list is complete.
            builder.ListFormat.RemoveNumbers();

            // Simple validation: count paragraphs that are part of a list.
            int listParagraphCount = 0;
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ListFormat.IsListItem)
                    listParagraphCount++;
            }

            // Output validation result to the console.
            Console.WriteLine($"Number of list paragraphs: {listParagraphCount}");

            // Save the document to a file.
            doc.Save("NineLevelList.docx");
        }
    }
}
