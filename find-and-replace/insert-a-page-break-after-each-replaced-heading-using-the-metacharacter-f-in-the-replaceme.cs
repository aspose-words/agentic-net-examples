using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");

        // Body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the first chapter.");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section A");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section A.");

        // Heading 2 again
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section B");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section B.");

        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Headings that need a page break after them.
        string[] headings = { "Chapter One", "Section A", "Section B" };
        int totalReplacements = 0;

        // Replace each heading with itself followed by a page break meta‑character (&m).
        foreach (string heading in headings)
        {
            // Use FindReplaceOptions with default settings.
            FindReplaceOptions options = new FindReplaceOptions();

            // Perform the replacement.
            int replaced = doc.Range.Replace(heading, heading + "&m", options);
            totalReplacements += replaced;
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}
