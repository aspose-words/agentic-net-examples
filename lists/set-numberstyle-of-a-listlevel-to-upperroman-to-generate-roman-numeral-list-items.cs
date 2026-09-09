using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level to use uppercase Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;
        // Use the default placeholder for the number format.
        level.NumberFormat = "\x0000.";

        // Use DocumentBuilder to write list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // Add several paragraphs that will be formatted as list items.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Item {i + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("RomanList.docx");
    }
}
