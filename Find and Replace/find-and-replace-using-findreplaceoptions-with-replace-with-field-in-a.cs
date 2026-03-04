using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options with a custom callback that inserts a field.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FieldReplacingHandler()
        };

        // Replace the placeholder "_Date_" with a DATE field.
        // The replacement string is empty because the actual insertion is performed in the callback.
        doc.Range.Replace("_Date_", string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

// Callback that replaces each match with a field.
class FieldReplacingHandler : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The document that contains the match is returned as DocumentBase; cast it to Document.
        Document doc = (Document)args.MatchNode.Document;

        // Create a DocumentBuilder for that document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder at the start of the node that contains the match.
        builder.MoveTo(args.MatchNode);

        // Insert the desired field. Example: a DATE field with a custom format.
        builder.InsertField("DATE", "\\@ \"MMMM d, yyyy\"");

        // Remove the original text that was matched.
        args.MatchNode.Remove();

        // Skip the default replacement because we have already performed the custom insertion.
        return ReplaceAction.Skip;
    }
}
