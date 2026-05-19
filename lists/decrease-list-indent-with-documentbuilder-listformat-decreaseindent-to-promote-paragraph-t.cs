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
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Increase the list level to create a sub‑list.
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 2.1");
        builder.Writeln("Subitem 2.2");

        // Decrease the list level to promote the next paragraph back to the higher level.
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 3");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("DecreaseIndent.docx");
    }
}
