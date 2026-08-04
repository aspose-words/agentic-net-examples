using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the starting number of the first list level to 5.
        // This means the first item will be numbered "5."
        list.ListLevels[0].StartAt = 5;

        // Use DocumentBuilder to add paragraphs that belong to the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Apply the list to subsequent paragraphs.

        // Add a few list items. They will be numbered 5, 6, 7, ...
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Remove list formatting from further paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Define an output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ListStartAtFive.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
