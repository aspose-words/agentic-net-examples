using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list and add two items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First item");
        builder.Writeln("Second item");

        // Retrieve the ListId of the list we just created.
        int listId = builder.CurrentParagraph.ListFormat.List.ListId;

        // Get the List object from the document by its ID.
        Aspose.Words.Lists.List listObj = doc.Lists.GetListByListId(listId);

        // Adjust level properties: change the first level to a bullet style,
        // set a custom font and size for the level's numbering.
        ListLevel level = listObj.ListLevels[0];

        // Instead of using ListNumberStyle (which may not be available in some versions),
        // set the NumberFormat directly to a bullet character.
        level.NumberFormat = "•";

        level.Font.Name = "Arial";
        level.Font.Size = 12;

        // Save the document to verify changes.
        string outputPath = "AdjustedList.docx";
        doc.Save(outputPath);

        // Indicate completion (no user interaction required).
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
