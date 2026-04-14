using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a numbered list with a few items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Build a nested list level.
        builder.ListFormat.ListIndent();
        builder.Writeln("Nested Item 1");
        builder.Writeln("Nested Item 2");
        builder.ListFormat.ListOutdent();

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Select all paragraphs that are part of a list.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                            .Cast<Paragraph>()
                            .Where(p => p.ListFormat.IsListItem);

        // Remove list formatting while preserving indentation.
        foreach (var paragraph in paragraphs)
        {
            paragraph.ListFormat.RemoveNumbers();
        }

        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "RemoveListFormatting.docx"));
    }
}
