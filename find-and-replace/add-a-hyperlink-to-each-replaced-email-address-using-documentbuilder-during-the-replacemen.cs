using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains a few e‑mail addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please contact john.doe@example.com for support.");
        builder.Writeln("You can also reach out to jane.smith@domain.org.");
        doc.Save("input.docx");

        // Load the document that we just created.
        Document loaded = new Document("input.docx");

        // Regular expression that matches a simple e‑mail address.
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}");

        // Set up replace options with a custom callback that inserts a hyperlink.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new EmailHyperlinkCallback();

        // Perform the replace operation. The callback will insert a hyperlink for each match.
        // We pass an empty replacement string because the callback handles the insertion.
        int replacedCount = loaded.Range.Replace(emailRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found for replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }

    // Callback that inserts a clickable mailto hyperlink for each e‑mail address found.
    private class EmailHyperlinkCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The document being processed.
            Document doc = (Document)args.MatchNode.Document;

            // Position a DocumentBuilder at the node that contains the match.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(args.MatchNode);

            // Insert a hyperlink where the e‑mail address was.
            string email = args.Match.Value;
            builder.InsertHyperlink(email, "mailto:" + email, false);

            // Remove the original text (the match) by replacing it with an empty string.
            args.Replacement = string.Empty;

            // Let the engine perform the replacement so the count is incremented.
            return ReplaceAction.Replace;
        }
    }
}
