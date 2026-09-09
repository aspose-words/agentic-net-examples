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

        // -------------------------------------------------
        // 1. Create the original list and apply it to the first section.
        // -------------------------------------------------
        List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
        // Example formatting for the first level.
        originalList.ListLevels[0].Font.Color = Color.Red;
        originalList.ListLevels[0].Alignment = ListLevelAlignment.Right;

        builder.Writeln("Original List starts below:");
        builder.ListFormat.List = originalList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // 2. Insert a new section break.
        // -------------------------------------------------
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // -------------------------------------------------
        // 3. Clone the original list, modify its start value, and apply it to the new section.
        // -------------------------------------------------
        List clonedList = doc.Lists.AddCopy(originalList);
        // Change the starting number for the first level and its color to differentiate.
        clonedList.ListLevels[0].StartAt = 10;
        clonedList.ListLevels[0].Font.Color = Color.Blue;

        builder.Writeln("Cloned List starts below:");
        builder.ListFormat.List = clonedList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // 4. Save the document.
        // -------------------------------------------------
        doc.Save("ListsCloneExample.docx");
    }
}
