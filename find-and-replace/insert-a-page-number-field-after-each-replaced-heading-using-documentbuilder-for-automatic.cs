using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;   // Needed for FindReplaceOptions

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three headings.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading One");
        builder.Writeln("Some content under heading one.");

        builder.Writeln("Heading Two");
        builder.Writeln("Some content under heading two.");

        builder.Writeln("Heading Three");
        builder.Writeln("Some content under heading three.");

        // Save the sample input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Replace the heading text.
        int replacedCount = loaded.Range.Replace("Heading", "New Heading", new FindReplaceOptions());
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one heading replacement.");

        // Insert a page number field after each replaced heading.
        foreach (Paragraph paragraph in loaded.GetChildNodes(NodeType.Paragraph, true))
        {
            // Trim to ignore trailing paragraph marks.
            string text = paragraph.GetText().TrimEnd('\r', '\a');
            if (text.StartsWith("New Heading", StringComparison.Ordinal))
            {
                // Position the builder at the current heading paragraph.
                DocumentBuilder paraBuilder = new DocumentBuilder(loaded);
                paraBuilder.MoveTo(paragraph);
                // Insert a new paragraph after the heading.
                paraBuilder.InsertParagraph();
                // Insert a PAGE field for automatic page numbering.
                paraBuilder.InsertField(FieldType.FieldPage, true);
            }
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
