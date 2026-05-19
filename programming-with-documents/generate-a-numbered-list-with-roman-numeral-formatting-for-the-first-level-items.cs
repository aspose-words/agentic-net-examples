using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the first level (index 0) to use uppercase Roman numerals.
        list.ListLevels[0].NumberStyle = NumberStyle.UppercaseRoman;

        // Apply the list to the builder so that subsequent paragraphs become list items.
        builder.ListFormat.List = list;

        // Add some list items. The numbering will appear as I., II., III., etc.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file in the same folder as the executable.
        doc.Save("RomanList.docx");
    }
}
