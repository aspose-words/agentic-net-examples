using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new multi‑level numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level to use upper‑case Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;
        // Use the placeholder \x0000 to insert the level number, followed by a period.
        level.NumberFormat = "\x0000.";

        // Use DocumentBuilder to add paragraphs that belong to the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Start the list.

        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        builder.ListFormat.RemoveNumbers(); // End the list.

        // Save the document.
        doc.Save("RomanList.docx");
    }
}
