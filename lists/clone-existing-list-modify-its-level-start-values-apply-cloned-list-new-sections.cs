using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Create the original list and add a few items in the first section.
        // -----------------------------------------------------------------
        List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
        // Example customization: make the first level number red.
        originalList.ListLevels[0].Font.Color = Color.Red;

        builder.Writeln("Original list starts below:");
        builder.ListFormat.List = originalList;   // Apply the list to subsequent paragraphs.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();       // End the list.

        // -----------------------------------------------------------------
        // 2. Start a new section.
        // -----------------------------------------------------------------
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // -----------------------------------------------------------------
        // 3. Clone the original list, modify its start values, and use it.
        // -----------------------------------------------------------------
        // AddCopy creates a new list based on the source list and adds it to the document.
        List clonedList = doc.Lists.AddCopy(originalList);

        // Change the starting number for the desired levels.
        // Here we set level 0 to start at 10 and level 1 to start at 5 as an example.
        clonedList.ListLevels[0].StartAt = 10;
        clonedList.ListLevels[1].StartAt = 5;

        builder.Writeln("Cloned list with modified start values:");
        builder.ListFormat.List = clonedList;     // Apply the cloned list.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();       // End the list.

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        doc.Save("ClonedList.docx");
    }
}
