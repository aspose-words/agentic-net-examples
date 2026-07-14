using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Heading {i}");
        }

        // Save the sample document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true,
            ReplacingCallback = new InsertPageNumberAfterHeading()
        };

        // Replace the word "Heading" with "Section" and invoke the callback.
        int replacedCount = loaded.Range.Replace("Heading", "Section", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one heading replacement.");

        // Update all fields (PAGE fields) so they display correct page numbers.
        loaded.UpdateFields();

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that inserts a PAGE field after each heading that is replaced.
    private class InsertPageNumberAfterHeading : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is inside a Run; its parent paragraph holds the heading text.
            Paragraph headingParagraph = args.MatchNode.ParentNode as Paragraph;
            if (headingParagraph != null)
            {
                // Insert a new empty paragraph right after the heading paragraph.
                Paragraph pageFieldParagraph = new Paragraph(headingParagraph.Document);
                headingParagraph.ParentNode.InsertAfter(pageFieldParagraph, headingParagraph);

                // Use a DocumentBuilder to place a PAGE field into the new paragraph.
                DocumentBuilder builder = new DocumentBuilder((Document)headingParagraph.Document);
                builder.MoveTo(pageFieldParagraph);
                builder.InsertField(FieldType.FieldPage, true);
            }

            // Replace the matched text with the new heading text.
            args.Replacement = "Section";
            return ReplaceAction.Replace;
        }
    }
}
