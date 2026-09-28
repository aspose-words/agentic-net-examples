using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Access the first level of the list (level 0).
        ListLevel level = list.ListLevels[0];

        // Position the number at the left margin.
        level.NumberPosition = 0;

        // Set a custom tab stop after the number to align the text.
        // Here we set it to 72 points (1 inch) from the left margin.
        level.TabPosition = 72;

        // Apply the list to several paragraphs.
        for (int i = 1; i <= 5; i++)
        {
            // Start a new list item.
            builder.ListFormat.List = list;
            builder.Writeln($"Item {i} with custom tab stop");
            // End the list formatting for the next paragraph.
            builder.ListFormat.RemoveNumbers();
        }

        // Save the document to disk.
        doc.Save("ListTabStop.docx");

        // Indicate completion.
        Console.WriteLine("Document created: ListTabStop.docx");
    }
}
