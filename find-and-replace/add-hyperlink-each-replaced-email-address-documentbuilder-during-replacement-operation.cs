using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document and write some sample text containing e‑mail addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("For assistance contact support@example.com or sales@example.org.");

        // Set up FindReplaceOptions with a custom callback that will replace each e‑mail with a hyperlink.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new EmailHyperlinkCallback();

        // Regular expression that matches e‑mail addresses.
        Regex emailRegex = new Regex(@"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b",
                                      RegexOptions.IgnoreCase);

        // Perform the replace operation. The replacement string is empty because the callback
        // inserts the hyperlink manually.
        doc.Range.Replace(emailRegex, string.Empty, options);

        // Save the resulting document.
        doc.Save("EmailHyperlinks.docx");
    }

    // Callback that is invoked for each regex match.
    private class EmailHyperlinkCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The matched e‑mail address.
            string email = args.Match.Value;

            // Remove the original text – the callback will insert the hyperlink instead.
            args.Replacement = string.Empty;

            // Create a DocumentBuilder positioned at the start of the node that contained the match.
            // args.MatchNode.Document returns DocumentBase, so cast to Document for the constructor.
            DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
            builder.MoveTo(args.MatchNode);

            // Apply typical hyperlink formatting.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert the hyperlink. The URL uses the mailto: scheme.
            builder.InsertHyperlink(email, "mailto:" + email, false);

            // Clear formatting so subsequent text is not affected.
            builder.Font.ClearFormatting();

            // Continue with the replace operation.
            return ReplaceAction.Replace;
        }
    }
}
