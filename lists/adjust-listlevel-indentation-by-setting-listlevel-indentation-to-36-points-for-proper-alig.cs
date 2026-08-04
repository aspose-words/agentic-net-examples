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

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multi‑level list based on the built‑in NumberDefault template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of each list level to 36 points (0.5 inch).
        // The ListLevel class does not have an Indentation property.
        // Instead, set NumberPosition, TextPosition and TabPosition to control the left indent.
        foreach (ListLevel level in list.ListLevels)
        {
            level.NumberPosition = 36.0; // Position of the number/bullet.
            level.TextPosition = 36.0;   // Position where the text starts.
            level.TabPosition = 36.0;    // Tab stop for the level.
        }

        // Apply the list to the builder and add a few items.
        builder.ListFormat.List = list;
        builder.Writeln("First level item");
        builder.ListFormat.ListIndent(); // Move to second level.
        builder.Writeln("Second level item");
        builder.ListFormat.ListIndent(); // Move to third level.
        builder.Writeln("Third level item");
        builder.ListFormat.ListOutdent(); // Back to second level.
        builder.Writeln("Another second level item");
        builder.ListFormat.RemoveNumbers(); // End the list.

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "Lists.AdjustIndentation.docx");
        doc.Save(outputPath);
    }
}
