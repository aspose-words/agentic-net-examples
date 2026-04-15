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

        // Add a multilevel list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the second level (index 1) to use lower‑case letters and a larger indent.
        ListLevel secondLevel = list.ListLevels[1];
        secondLevel.NumberStyle = NumberStyle.LowercaseLetter; // a., b., c., …
        secondLevel.NumberPosition = -36;                       // Move the number slightly left.
        secondLevel.TextPosition = 72;                          // Indent the text after the number.

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First‑level items.
        builder.Writeln("First level item 1");
        builder.Writeln("First level item 2");

        // Increase the list level to the second level.
        builder.ListFormat.ListIndent();

        // Second‑level items (will be a., b., …).
        builder.Writeln("Second level item a");
        builder.Writeln("Second level item b");
        builder.Writeln("Second level item c");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // More first‑level items.
        builder.Writeln("First level item 3");
        builder.Writeln("First level item 4");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }
}
