using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // 1. Create an original list and add a few items.
            // -----------------------------------------------------------------
            List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            // Example formatting for the first level.
            originalList.ListLevels[0].Font.Color = Color.Red;
            originalList.ListLevels[0].Alignment = ListLevelAlignment.Right;

            builder.Writeln("Original list starts below:");
            builder.ListFormat.List = originalList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // -----------------------------------------------------------------
            // 2. Clone the existing list, modify its start values.
            // -----------------------------------------------------------------
            List clonedList = doc.Lists.AddCopy(originalList);
            // Change the starting number for the first level to 10.
            clonedList.ListLevels[0].StartAt = 10;
            // Optionally change the start value for a second level.
            if (clonedList.ListLevels.Count > 1)
                clonedList.ListLevels[1].StartAt = 5;

            // -----------------------------------------------------------------
            // 3. Add a new section and apply the cloned list there.
            // -----------------------------------------------------------------
            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln("Cloned list starts below (starts at 10):");
            builder.ListFormat.List = clonedList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // -----------------------------------------------------------------
            // 4. Save the document.
            // -----------------------------------------------------------------
            doc.Save("ClonedList.docx");
        }
    }
}
