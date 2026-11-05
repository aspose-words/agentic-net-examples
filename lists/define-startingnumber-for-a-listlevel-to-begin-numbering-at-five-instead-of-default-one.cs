using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the starting number of the first level to 5.
        ListLevel level0 = list.ListLevels[0];
        level0.StartAt = 5; // Correct property to define the starting number.

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Validate that three paragraphs are list items.
        int listItemCount = 0;
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            if (para.ListFormat.IsListItem)
                listItemCount++;
        }

        // Output validation result.
        Console.WriteLine($"List items written: {listItemCount}");

        // Save the document to a file.
        doc.Save("ListStartingNumber.docx");
    }
}
