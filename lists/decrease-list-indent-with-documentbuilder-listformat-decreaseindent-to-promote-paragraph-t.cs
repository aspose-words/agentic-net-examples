using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // First level item.
        builder.Writeln("First level item");

        // Increase the list level – this creates a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Second level item (indented)");

        // Decrease the list level – promotes the next paragraph back to the first level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("First level item after outdent");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = "Lists.DecreaseIndent.docx";
        doc.Save(outputPath);
    }
}
