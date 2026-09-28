using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading One");
        builder.Writeln("Content under first heading.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading Two");
        builder.Writeln("Content under second heading.");

        // Set up the replacement callback.
        var callback = new HeadingReplaceCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("Heading", "Section", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    private class HeadingReplaceCallback : IReplacingCallback
    {
        // Track paragraphs that have already received a page number field.
        private readonly HashSet<Paragraph> _processedParagraphs = new HashSet<Paragraph>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace the matched text.
            args.Replacement = "Section";

            // Get the paragraph that contains the match.
            Node matchNode = args.MatchNode;
            if (matchNode?.ParentNode is Paragraph paragraph && !_processedParagraphs.Contains(paragraph))
            {
                // Insert a new paragraph after the heading.
                DocumentBuilder builder = new DocumentBuilder((Document)paragraph.Document, new DocumentBuilderOptions());
                builder.MoveTo(paragraph);
                builder.InsertParagraph();

                // Insert a PAGE field for automatic numbering.
                builder.InsertField(FieldType.FieldPage, true);

                _processedParagraphs.Add(paragraph);
            }

            return ReplaceAction.Replace;
        }
    }
}
