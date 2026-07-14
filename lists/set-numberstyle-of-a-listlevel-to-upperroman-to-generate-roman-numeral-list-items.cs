using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new list based on the default numbered template (contains 9 levels).
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level (level 0) to use upper‑case Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;
        // Define the format string – the placeholder \x0000 will be replaced by the Roman numeral.
        level.NumberFormat = "\x0000.";

        // Use DocumentBuilder to write paragraphs that belong to the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Start the list.

        // Add a few list items.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Item {i + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("ListUpperRoman.docx");
    }
}
