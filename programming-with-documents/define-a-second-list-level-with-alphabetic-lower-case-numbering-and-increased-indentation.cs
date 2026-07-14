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

        // Start a default numbered list (level 0 uses Arabic numbers).
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Increase the list level to create a second level.
        builder.ListFormat.ListIndent();

        // Retrieve the underlying List object to customize the second level.
        List list = builder.ListFormat.List;
        ListLevel secondLevel = list.ListLevels[1]; // Index 1 = second level (0‑based).

        // Ensure the second level uses lower‑case alphabetic numbering.
        secondLevel.NumberStyle = NumberStyle.LowercaseLetter;

        // Increase indentation for the second level.
        // NumberPosition: position of the list label (negative moves it left of the margin).
        // TextPosition: start position of the paragraph text.
        secondLevel.NumberPosition = -36;   // Move the label slightly left.
        secondLevel.TextPosition = 144;     // Indent the text further to the right.
        secondLevel.TabPosition = 144;      // Align tab stops with the text indent.

        // Add items to the second level.
        builder.Writeln("Second level a");
        builder.Writeln("Second level b");
        builder.Writeln("Second level c");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // Add more first‑level items.
        builder.Writeln("First level item 3");
        builder.Writeln("First level item 4");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "SecondListLevel.docx");
        doc.Save(outputPath);
    }
}
