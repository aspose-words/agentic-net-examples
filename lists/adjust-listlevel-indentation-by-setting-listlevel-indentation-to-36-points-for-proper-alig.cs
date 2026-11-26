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

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the indentation of the first list level to 36 points.
        // In Aspose.Words the indentation is controlled by NumberPosition (position of the number)
        // and TextPosition (position of the list text). Setting both to 36 points aligns the list.
        list.ListLevels[0].NumberPosition = 36;
        list.ListLevels[0].TextPosition = 36;

        // Apply the list to the builder and set the level.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0;

        // Write list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("ListIndentation.docx");
    }
}
