using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FieldReplacingCallback()
        };

        // Replace the placeholder "_FullName_" with a MERGEFIELD.
        // The replacement string is empty because the callback inserts the field manually.
        doc.Range.Replace("_FullName_", string.Empty, options);

        // Save the updated document.
        doc.Save("Output.docx");
    }

    // Callback that replaces the matched text with a field.
    private class FieldReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The node that contains the match is a Run (or another text node).
            // Cast the document to Document because DocumentBuilder expects a Document, not DocumentBase.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
            builder.MoveTo(args.MatchNode);

            // Insert the desired field (e.g., a MERGEFIELD named FullName).
            builder.InsertField("MERGEFIELD FullName", "«FullName»");

            // Remove the original placeholder text node.
            args.MatchNode.Remove();

            // Skip the default replacement because we have already performed the insertion.
            return ReplaceAction.Skip;
        }
    }
}
