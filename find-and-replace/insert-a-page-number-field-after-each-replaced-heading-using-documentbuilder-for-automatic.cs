using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");
        builder.Writeln("Content under heading 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2");
        builder.Writeln("Content under heading 2.");

        // Save the initial document (optional, for inspection).
        doc.Save("input.docx");

        // Set up find‑and‑replace with a callback that inserts a PAGE field after each heading.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new HeadingReplaceCallback();

        // Replace the word "Heading" with "New Heading". The callback will add a PAGE field after each heading.
        int replacedCount = doc.Range.Replace("Heading", "New Heading", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one heading replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that runs for each match found during replacement.
    private class HeadingReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the matched text.
            Paragraph headingParagraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
            if (headingParagraph != null)
            {
                // Use a DocumentBuilder tied to the same document.
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);

                // Move the builder to the heading paragraph.
                builder.MoveTo(headingParagraph);

                // Insert a new empty paragraph after the heading.
                builder.Writeln();

                // Insert the PAGE field for automatic page numbering.
                builder.InsertField(FieldType.FieldPage, true);
            }

            // Perform the default replacement (replace the matched text with the new string).
            return ReplaceAction.Replace;
        }
    }
}
