using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level (index 0) to use uppercase Roman numerals.
        ListLevel firstLevel = list.ListLevels[0];
        firstLevel.NumberStyle = NumberStyle.UppercaseRoman;
        // Optional: define the format string if you want a trailing period.
        // firstLevel.NumberFormat = "%1.";

        // Use DocumentBuilder to insert list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        // Add several items to the list.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RomanNumberedList.docx");
        doc.Save(outputPath);
    }
}
