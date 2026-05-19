using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing email addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please contact us at support@example.com or sales@example.org for more information.");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Regular expression that matches email addresses.
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}");

        // Set up find‑replace options with a custom callback that inserts a hyperlink.
        FindReplaceOptions options = new FindReplaceOptions(new EmailHyperlinkCallback());

        // Perform the replace operation. The callback will replace each email with a mailto hyperlink.
        int replacedCount = loaded.Range.Replace(emailRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found to replace.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that replaces each email address with a mailto hyperlink.
    private class EmailHyperlinkCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The matched email address.
            string email = args.Match.Value;

            // Remove the original text.
            args.Replacement = string.Empty;

            // Insert a hyperlink at the position of the match.
            // Cast the DocumentBase to Document because DocumentBuilder expects a Document instance.
            DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);
            cb.MoveTo(args.MatchNode);
            cb.InsertHyperlink(email, "mailto:" + email, false);

            // Continue with the replacement (which removes the original email text).
            return ReplaceAction.Replace;
        }
    }
}
