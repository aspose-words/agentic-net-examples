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

        // Store the ListId for later retrieval.
        int listId = list.ListId;

        // Retrieve the same list using its identifier.
        List retrievedList = doc.Lists.GetListByListId(listId);

        // If the list exists, modify its first level formatting.
        if (retrievedList != null)
        {
            ListLevel level = retrievedList.ListLevels[0];
            level.Font.Color = Color.Blue;          // Change bullet/number color.
            level.Font.Size = 16;                    // Increase font size.
            level.NumberStyle = NumberStyle.UppercaseRoman; // Use uppercase Roman numerals.
            level.StartAt = 5;                       // Start numbering at 5.
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
