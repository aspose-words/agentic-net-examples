using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new list based on the default numbered template (contains 9 levels).
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list to use upper‑case Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;   // I, II, III, ...
        level.NumberFormat = "\x0000";                    // Placeholder for the level's number.
        level.Font.Size = 12;
        level.Font.Color = Color.Black;

        // Use a DocumentBuilder to insert paragraphs that will be formatted as list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Apply the custom list to subsequent paragraphs.

        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Item {i + 1}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file in the current directory.
        string outputPath = "RomanList.docx";
        doc.Save(outputPath);
    }
}
