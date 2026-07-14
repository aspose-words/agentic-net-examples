using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the starting number of the first list level to 5.
        // This will cause the first item to be numbered "5."
        list.ListLevels[0].StartAt = 5;

        // Use DocumentBuilder to add paragraphs that belong to the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Apply the list to subsequent paragraphs.

        // Add a few list items to demonstrate the custom start number.
        builder.Writeln("Item 1 (should be numbered 5)");
        builder.Writeln("Item 2 (should be numbered 6)");
        builder.Writeln("Item 3 (should be numbered 7)");

        // Remove list formatting from further paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output file.
        string outputPath = "Lists.StartAt.docx";
        doc.Save(outputPath);
    }
}
