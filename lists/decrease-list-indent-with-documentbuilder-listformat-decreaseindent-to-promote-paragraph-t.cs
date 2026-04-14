using System;
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
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item at level 0");

        // Increase the list level to create a sub‑list (level 1).
        builder.ListFormat.ListIndent();
        builder.Writeln("Item at level 1");

        // Increase again to level 2.
        builder.ListFormat.ListIndent();
        builder.Writeln("Item at level 2");

        // Decrease the list level back to level 1.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Back to level 1");

        // Decrease once more to return to the top level (level 0).
        builder.ListFormat.ListOutdent();
        builder.Writeln("Back to level 0");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("DecreaseIndent.docx");
    }
}
