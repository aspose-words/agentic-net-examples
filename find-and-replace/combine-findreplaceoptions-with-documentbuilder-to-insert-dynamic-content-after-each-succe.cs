using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing the placeholder "_Name_".
        builder.Writeln("Dear _Name_,");
        builder.Writeln("Welcome to the company.");
        builder.Writeln("Your colleague _Name_ will greet you.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertAfterCallback()
        };

        // Perform the replacement. The callback will insert additional content after each match.
        int replacementCount = doc.Range.Replace("_Name_", "John", options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }

    // Callback that runs for each match found during the replace operation.
    private class InsertAfterCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ensure the match node belongs to a paragraph.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // Create a builder for the document that contains the matched paragraph.
                DocumentBuilder cb = new DocumentBuilder((Document)paragraph.Document);

                // Move the builder to the matched paragraph.
                cb.MoveTo(paragraph);

                // Insert a new empty paragraph after the current one.
                cb.InsertParagraph();

                // Write dynamic content into the newly inserted paragraph.
                cb.Writeln($"[Inserted after \"{args.Match.Value}\" replacement]");
            }

            // Continue with the normal replacement (the replacement text is supplied by the caller).
            return ReplaceAction.Replace;
        }
    }
}
