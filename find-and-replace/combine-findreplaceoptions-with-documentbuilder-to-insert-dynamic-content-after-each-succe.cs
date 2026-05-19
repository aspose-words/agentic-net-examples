using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line with placeholder.");
        builder.Writeln("Second line with placeholder.");
        builder.Writeln("Third line without the keyword.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertAfterReplacementCallback(doc)
        };

        // Replace the word "placeholder" and let the callback insert dynamic content.
        int replacedCount = doc.Range.Replace("placeholder", "replaced", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that inserts a new paragraph with dynamic content after each match.
    private class InsertAfterReplacementCallback : IReplacingCallback
    {
        private readonly Document _document;
        private int _insertionIndex = 0;

        public InsertAfterReplacementCallback(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Increment a counter to generate unique dynamic text.
            _insertionIndex++;
            string dynamicText = $"[Inserted {_insertionIndex} at {DateTime.Now:HH:mm:ss}]";

            // The match is located inside a Run; its parent is a Paragraph.
            Paragraph paragraph = args.MatchNode.ParentNode as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Use a new DocumentBuilder to insert a paragraph after the current one.
            DocumentBuilder cb = new DocumentBuilder(_document);
            cb.MoveTo(paragraph);
            cb.Writeln(dynamicText);

            // Optionally modify the replacement text.
            args.Replacement = "replaced";

            return ReplaceAction.Replace;
        }
    }
}
