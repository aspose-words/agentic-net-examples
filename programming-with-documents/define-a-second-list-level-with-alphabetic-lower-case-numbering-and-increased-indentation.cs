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

        // Add a heading for the list.
        builder.Writeln("Sample list with a customized second level:");

        // Start a default numbered list (level 0 uses Arabic numbers).
        builder.ListFormat.ApplyNumberDefault();

        // First item at level 0.
        builder.Writeln("First level item 1");

        // Retrieve the underlying List object to modify its second level (index 1).
        List list = builder.ListFormat.List;
        ListLevel secondLevel = list.ListLevels[1];

        // Ensure the second level uses lower‑case letters (a., b., …) and increase its indentation.
        secondLevel.NumberStyle = NumberStyle.LowercaseLetter;
        // NumberPosition is the position of the list label; a negative value moves it left.
        secondLevel.NumberPosition = -18;   // move the label slightly left of the text.
        // TextPosition is the start of the paragraph text; increase it for a deeper indent.
        secondLevel.TextPosition = 36;      // indent the text further to the right.

        // Increase the list level to the second level.
        builder.ListFormat.ListIndent();

        // Items at the second level will be numbered with lower‑case letters.
        builder.Writeln("Second level item a");
        builder.Writeln("Second level item b");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // Another first‑level item.
        builder.Writeln("First level item 2");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ListWithCustomSecondLevel.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document.
        doc.Save(outputPath);
    }
}
