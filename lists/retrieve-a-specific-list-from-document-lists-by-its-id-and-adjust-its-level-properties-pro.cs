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

        // Use DocumentBuilder to add paragraphs that belong to the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Store the ListId of the created list.
        int listId = list.ListId;

        // Retrieve the same list by its identifier.
        List retrievedList = doc.Lists.GetListByListId(listId);
        if (retrievedList != null)
        {
            // Modify properties of the first level of the list.
            retrievedList.ListLevels[0].Font.Color = Color.Blue; // Change bullet/number color.
            retrievedList.ListLevels[0].StartAt = 10;            // Change the starting number.
        }

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
