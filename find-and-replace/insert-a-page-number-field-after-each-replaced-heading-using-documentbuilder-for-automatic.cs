using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

public class InsertPageNumberAfterReplacedHeading
{
    public static void Main()
    {
        // Create a new document and a builder for constructing its content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample headings that we will replace later.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.Writeln("Some introductory text.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Two");
        builder.Writeln("More content follows.");

        // Prepare find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertPageNumberCallback(doc);

        // Replace each heading text (exact match) with a new title.
        // The callback will insert a PAGE field after each replaced heading.
        int replacements = doc.Range.Replace("Chapter", "Section", options);

        // Ensure that at least one replacement was performed.
        if (replacements == 0)
            throw new InvalidOperationException("No headings were replaced.");

        // Update all fields so that PAGE fields show correct numbers.
        doc.UpdateFields();

        // Save the resulting document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}. Replacements made: {replacements}");
    }

    // Callback that inserts a PAGE field after each matched heading.
    private class InsertPageNumberCallback : IReplacingCallback
    {
        private readonly Document _document;

        public InsertPageNumberCallback(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Keep the original matched text (or modify it as needed).
            args.Replacement = args.Match.Value;

            // The node that contains the start of the match.
            // Its parent is the paragraph that holds the heading.
            Paragraph headingParagraph = args.MatchNode.ParentNode as Paragraph;
            if (headingParagraph == null)
                return ReplaceAction.Skip;

            // Create a new paragraph that will contain the PAGE field.
            Paragraph pageNumberParagraph = new Paragraph(_document);
            // Insert the new paragraph immediately after the heading paragraph.
            headingParagraph.ParentNode.InsertAfter(pageNumberParagraph, headingParagraph);

            // Use a DocumentBuilder to place a PAGE field inside the new paragraph.
            DocumentBuilder fieldBuilder = new DocumentBuilder(_document);
            fieldBuilder.MoveTo(pageNumberParagraph);
            fieldBuilder.InsertField(FieldType.FieldPage, true);

            return ReplaceAction.Replace;
        }
    }
}
