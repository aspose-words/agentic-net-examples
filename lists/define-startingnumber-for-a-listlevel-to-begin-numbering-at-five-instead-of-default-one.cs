using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list based on the built‑in NumberDefault template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the starting number of the first list level to 5.
        // This means the first item will be numbered "5.", then "6.", etc.
        list.ListLevels[0].StartAt = 5;

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;

        // Insert a few list items to demonstrate the custom start number.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Remove list formatting from the builder.
        builder.ListFormat.RemoveNumbers();

        // Define an output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "ListStartAtFive.docx");
        doc.Save(outputPath);
    }
}
