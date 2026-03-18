using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a numbered list based on the default list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Get the first list level (level 0) and set its numbering style to Uppercase Roman.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;

        // Define the format string so the Roman numeral appears followed by a period.
        level.NumberFormat = "\x0000.";

        // Use DocumentBuilder to insert list items that will use the configured list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Apply the list to the current paragraph.

        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Item {i + 1}");
        }

        // Save the document to disk.
        doc.Save("RomanList.docx");
    }
}
