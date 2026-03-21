using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a simple document with a numbered list.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First item");
        builder.Writeln("Second item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Retrieve the first list in the document.
        if (doc.Lists.Count == 0)
        {
            Console.WriteLine("No lists were found in the document.");
            return;
        }

        List list = doc.Lists[0]; // The first (and only) list.

        // Modify properties of the first level (index 0) of the list.
        ListLevel level = list.ListLevels[0];

        // Change the font color of the list label.
        level.Font.Color = Color.Green;

        // Set the alignment of the list number/bullet.
        level.Alignment = ListLevelAlignment.Left;

        // Change the starting number for this level.
        level.StartAt = 5;

        // Adjust indentation properties.
        level.NumberPosition = -18;   // Position of the number/bullet.
        level.TextPosition = 36;      // Position of the text after the number/bullet.
        level.TabPosition = 36;       // Tab stop after the number/bullet.

        // Determine output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
