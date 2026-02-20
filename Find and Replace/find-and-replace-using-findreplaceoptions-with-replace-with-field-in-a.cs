using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

namespace FindReplaceWithFieldExample
{
    // Custom callback that replaces each match with a DATE field.
    class InsertDateFieldHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The Document property of MatchNode returns DocumentBase; cast it to Document.
            Document doc = (Document)args.MatchNode.Document;
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the start of the match.
            // The match may span multiple runs; moving to the match node positions the cursor correctly.
            builder.MoveTo(args.MatchNode);

            // Insert a DATE field at the match location.
            builder.InsertField(FieldType.FieldDate, true);

            // Skip the default text replacement because we have already inserted the field.
            return ReplaceAction.Skip;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("Input.docx");

            // Set up find/replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertDateFieldHandler();

            // Replace all occurrences of the placeholder text with a DATE field.
            // The replacement string is empty because the callback handles insertion.
            doc.Range.Replace(new Regex(@"\[DATE\]"), string.Empty, options);

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
