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

        // Add a new list to the document (using a default numbered list as a base).
        List romanList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list to use uppercase Roman numerals.
        ListLevel level = romanList.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman; // Roman numeral style.
        level.NumberFormat = "%1.";                     // Append a period after the numeral.
        level.StartAt = 1;                              // Start numbering at I.
        level.Alignment = ListLevelAlignment.Left;

        // Apply the list to the builder.
        builder.ListFormat.List = romanList;

        // Insert list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        string outputPath = "RomanNumberedList.docx";
        doc.Save(outputPath);
    }
}
