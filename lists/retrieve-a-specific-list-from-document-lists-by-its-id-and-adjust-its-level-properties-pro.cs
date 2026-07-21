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

        // Add a default numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Use DocumentBuilder to add a few list items.
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

        // Adjust properties of the first level of the retrieved list.
        if (retrievedList != null)
        {
            ListLevel level0 = retrievedList.ListLevels[0];
            level0.StartAt = 10;                     // Change the starting number.
            level0.Font.Color = Color.Blue;          // Set the bullet/number color.
            level0.NumberStyle = NumberStyle.Bullet; // Switch to a bullet style.
        }

        // Save the modified document.
        doc.Save("RetrieveAndModifyList.docx");
    }
}
