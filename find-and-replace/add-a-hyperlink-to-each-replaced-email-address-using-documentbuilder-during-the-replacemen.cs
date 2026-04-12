using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing e‑mail addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please contact us at support@example.com for assistance.");
        builder.Writeln("You can also reach sales@example.org or info@example.net.");

        // Define a regular expression that matches e‑mail addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.IgnoreCase);

        // Set up FindReplaceOptions with a custom callback that inserts a hyperlink.
        FindReplaceOptions options = new FindReplaceOptions(new EmailHyperlinkCallback());

        // Perform the replace operation. The callback will handle the actual insertion.
        int replacedCount = doc.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No e‑mail addresses were found for replacement.");

        // Save the resulting document.
        doc.Save("HyperlinkEmails.docx");
    }

    // Callback that replaces each e‑mail match with a clickable mailto: hyperlink.
    private class EmailHyperlinkCallback : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Build the mailto URL.
            string mailto = "mailto:" + args.Match.Value;

            // Insert the hyperlink at the position of the match.
            // Cast to Document because DocumentBuilder expects a Document, not DocumentBase.
            DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);
            cb.MoveTo(args.MatchNode);
            cb.InsertHyperlink(args.Match.Value, mailto, false);

            // Delete the original matched text by replacing it with an empty string.
            args.Replacement = string.Empty;

            // Let the replace engine perform the deletion.
            return ReplaceAction.Replace;
        }
    }
}
