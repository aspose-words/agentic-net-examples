using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // First level list item.
        builder.Writeln("Item 1");

        // Increase indent to create a second‑level item.
        builder.ListFormat.ListLevelNumber++;
        builder.Writeln("Item 1.1");

        // Increase indent again for a third‑level item.
        builder.ListFormat.ListLevelNumber++;
        builder.Writeln("Item 1.1.1");

        // Decrease indent back to second level.
        builder.ListFormat.ListLevelNumber--;
        builder.Writeln("Item 1.2");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("ListIndent.docx");
    }
}
