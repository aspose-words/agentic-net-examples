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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1 - level 0");
        builder.Writeln("Item 2 - level 0");

        // Increase the list level (indent) to create a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 2.1 - level 1");
        builder.Writeln("Item 2.2 - level 1");

        // Increase the list level again for a deeper sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item 2.2.1 - level 2");

        // Decrease the list level (outdent) back to the previous level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 2.3 - back to level 1");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ListIndentExample.docx");
        doc.Save(outputPath);
    }
}
