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

        // Create a new list based on the default numbered list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list to use upper‑case Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;   // Correct enum value for Roman numerals
        level.NumberFormat = "%1.";                       // Append a period after the number
        level.Alignment = ListLevelAlignment.Left;        // Align left using ListLevelAlignment
        level.StartAt = 1;                                // Start numbering at I

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Add first‑level items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output file.
        string outputPath = "RomanNumberedList.docx";
        doc.Save(outputPath);
    }
}
