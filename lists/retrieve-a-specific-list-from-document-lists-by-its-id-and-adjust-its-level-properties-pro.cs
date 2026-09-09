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

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Use DocumentBuilder to add a few list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Store the unique identifier of the created list.
        int listId = list.ListId;

        // Retrieve the same list from the collection by its identifier.
        List retrievedList = doc.Lists.GetListByListId(listId);
        if (retrievedList != null)
        {
            // Modify properties of the first level of the list.
            // Change the font color to blue and set the starting number to 10.
            retrievedList.ListLevels[0].Font.Color = Color.Blue;
            retrievedList.ListLevels[0].StartAt = 10;
        }

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
