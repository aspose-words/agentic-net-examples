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

        // Start a default numbered list.
        // Level 0 → Arabic numbers (1., 2., …)
        // Level 1 → Lower‑case letters (a., b., …)
        builder.ListFormat.ApplyNumberDefault();

        // First‑level list item.
        builder.Writeln("First level item 1");

        // Move to the second list level.
        builder.ListFormat.ListIndent();

        // Increase indentation for the second level.
        // NumberPosition defines where the list symbol is placed.
        // TextPosition defines where the paragraph text starts.
        builder.ListFormat.ListLevel.NumberPosition = -36; // number left of the text
        builder.ListFormat.ListLevel.TextPosition = 144;   // text starts further right

        // Second‑level items (automatically use lower‑case letters).
        builder.Writeln("Second level item a");
        builder.Writeln("Second level item b");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // Another first‑level item.
        builder.Writeln("First level item 2");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "SecondLevelList.docx");
        doc.Save(outputPath);
    }
}
