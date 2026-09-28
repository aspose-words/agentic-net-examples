using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");

        // Add a second top‑level item.
        builder.Writeln("Item 2");

        // Increase indent to create a sub‑list by setting the list level number.
        builder.ListFormat.ListLevelNumber = 1; // level 2 (zero‑based)
        builder.Writeln("Item 2.1");
        builder.Writeln("Item 2.2");

        // Decrease indent to promote back to the higher level.
        builder.ListFormat.ListLevelNumber = 0; // back to top level
        builder.Writeln("Item 3");

        // Save the document to disk.
        doc.Save("ListIndent.docx");
    }
}
