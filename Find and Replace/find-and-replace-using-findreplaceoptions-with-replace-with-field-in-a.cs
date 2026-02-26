using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new FieldReplacingCallback()
        };

        // Replace the placeholder text with a field.
        doc.Range.Replace("_FieldPlaceholder_", string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

// Callback that replaces each match with a DATE field.
class FieldReplacingCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The match node is a Run that contains the placeholder text.
        // Insert the field at the position of the match.
        DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
        builder.MoveTo(args.MatchNode);
        builder.InsertField("DATE", "\\@ \"MMMM d, yyyy\"");

        // Remove the original placeholder text.
        args.MatchNode.Remove();

        // Skip the default text replacement.
        return ReplaceAction.Skip;
    }
}
