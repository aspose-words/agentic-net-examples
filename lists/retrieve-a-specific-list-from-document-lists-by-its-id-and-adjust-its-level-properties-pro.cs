using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Use the list for a couple of paragraphs so the list is actually stored in the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Retrieve the same list by its unique identifier.
        int listId = list.ListId;
        List retrievedList = doc.Lists.GetListByListId(listId);

        // Adjust properties of the first level of the retrieved list.
        if (retrievedList != null)
        {
            ListLevel level0 = retrievedList.ListLevels[0];
            level0.Font.Color = Color.Blue;      // Change the bullet/number color.
            level0.StartAt = 5;                  // Start numbering at 5.
            level0.Alignment = ListLevelAlignment.Right; // Align the number to the right.
        }

        // Save the modified document.
        doc.Save("ListAdjusted.docx");
    }
}
