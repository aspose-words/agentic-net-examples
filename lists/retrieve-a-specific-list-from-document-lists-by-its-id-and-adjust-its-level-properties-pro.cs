using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        int listId = list.ListId; // Store the list ID for later retrieval.

        // Apply the list to the builder.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0; // Top level.

        // Add initial list items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // End the first list.
        builder.ListFormat.RemoveNumbers();

        // Retrieve the same list by its ID.
        List retrievedList = doc.Lists.GetListByListId(listId);

        // Adjust level properties of the retrieved list.
        ListLevel level0 = retrievedList.ListLevels[0];
        level0.NumberStyle = NumberStyle.UppercaseRoman;          // Change numbering to I, II, III...
        level0.Alignment = ListLevelAlignment.Right;              // Align numbers to the right.

        // Use the modified list again.
        builder.ListFormat.List = retrievedList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Modified Item 1");
        builder.Writeln("Modified Item 2");
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("Result.docx");
    }
}
