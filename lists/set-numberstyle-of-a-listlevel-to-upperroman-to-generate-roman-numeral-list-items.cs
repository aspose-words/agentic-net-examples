using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a list based on the built‑in numbered template.
        List romanList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list to use upper‑case Roman numerals.
        ListLevel level = romanList.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;
        // Use the default number format placeholder for the current level.
        level.NumberFormat = "\x0000.";

        // Use a DocumentBuilder to add paragraphs that will be formatted by the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = romanList;

        // Add several items – they will be numbered I., II., III., etc.
        builder.Writeln("First Roman item");
        builder.Writeln("Second Roman item");
        builder.Writeln("Third Roman item");
        builder.Writeln("Fourth Roman item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        string outputPath = "RomanList.docx";
        doc.Save(outputPath);
    }
}
