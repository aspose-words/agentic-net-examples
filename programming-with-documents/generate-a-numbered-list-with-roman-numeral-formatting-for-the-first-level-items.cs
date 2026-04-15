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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default template.
        List romanList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the first level (index 0) to use uppercase Roman numerals.
        romanList.ListLevels[0].NumberStyle = NumberStyle.UppercaseRoman;

        // Apply the list to the builder so subsequent paragraphs become list items.
        builder.ListFormat.List = romanList;

        // Add several list items.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Item {i}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "RomanNumberedList.docx");
        doc.Save(outputPath);
    }
}
