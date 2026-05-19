using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the current list (optional, just for clarity).
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph between two lists.
        builder.Writeln("Normal paragraph");

        // Start another numbered list and add two more items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 4");
        builder.Writeln("Item 5");

        // Retrieve all paragraphs in the document.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .ToList();

        // Remove list formatting from the second and third list items (preserving indentation).
        // These are the paragraphs at index 1 and 2 among the list items.
        foreach (var para in paragraphs
                             .Where(p => p.ListFormat.IsListItem)
                             .Skip(1)   // skip the first list item
                             .Take(2))  // take the next two items
        {
            // This call removes the bullet/number and sets the list level to zero,
            // while keeping any existing indentation of the paragraph.
            para.ListFormat.RemoveNumbers();
        }

        // Save the resulting document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RemoveListFormatting.docx");
        doc.Save(outputPath);
    }
}
