using System;
using Aspose.Words;
using Aspose.Words.Lists;

class RomanNumberedListExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list based on the default numbered template.
        List romanList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level (level 0) to use uppercase Roman numerals.
        ListLevel level0 = romanList.ListLevels[0];
        level0.NumberStyle = NumberStyle.UppercaseRoman;   // I, II, III, ...
        level0.NumberFormat = "\x0000.";                  // Append a period after the number.

        // Apply the customized list to the builder.
        builder.ListFormat.List = romanList;

        // Add first‑level items – they will be formatted as Roman numerals.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("RomanNumberedList.docx");
    }
}
