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

        // Add a numbered list to the document using a predefined template.
        List originalList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add a few paragraphs and apply the list to them.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        // Apply the list to the last three paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ListFormat.List = originalList;
            para.ListFormat.ListLevelNumber = 0; // first level
        }

        // Retrieve the ListId of the created list.
        int listId = originalList.ListId;

        // Use GetListByListId to obtain the same list from the collection.
        List retrievedList = doc.Lists.GetListByListId(listId);
        if (retrievedList == null)
        {
            Console.WriteLine("List with ID {0} not found.", listId);
            return;
        }

        // Adjust properties of the first level of the retrieved list.
        // For example, change the font color, start number, and number style.
        ListLevel level0 = retrievedList.ListLevels[0];
        level0.Font.Color = Color.Blue;          // Change bullet/number color.
        level0.StartAt = 10;                     // Start numbering at 10.
        level0.NumberStyle = NumberStyle.UppercaseRoman; // Use uppercase Roman numerals.

        // Save the document to verify the changes.
        doc.Save("Lists_RetrieveAndModify.docx");
    }
}
