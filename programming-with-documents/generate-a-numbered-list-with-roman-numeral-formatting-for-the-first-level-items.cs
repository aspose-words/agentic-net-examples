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

        // Create a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list to use uppercase Roman numerals.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberStyle = NumberStyle.UppercaseRoman;
        // Optional: define the format pattern (e.g., "I.", "II.", ...).
        level0.NumberFormat = "%1.";

        // Use DocumentBuilder to add paragraphs that belong to the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;

        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "RomanNumberedList.docx");
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
