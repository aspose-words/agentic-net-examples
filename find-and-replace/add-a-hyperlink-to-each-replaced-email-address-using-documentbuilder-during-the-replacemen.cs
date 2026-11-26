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
        builder.Writeln("Please contact us at support@example.com for assistance.");
        builder.Writeln("You can also reach sales@example.org for inquiries.");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Regex pattern that matches email addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.Compiled);

        // Set up find‑replace options with a custom callback.
        EmailHyperlinkCallback callback = new EmailHyperlinkCallback();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement. The replacement string is ignored because the callback handles insertion.
        loaded.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one email address was replaced.
        if (callback.Replacements == 0)
            throw new InvalidOperationException("No email addresses were found for replacement.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}

// Callback that replaces each found email address with a hyperlink.
public class EmailHyperlinkCallback : IReplacingCallback
{
    // Number of replacements performed.
    public int Replacements { get; private set; }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The matched email address.
        string email = args.Match.Value;

        // Obtain the Document that owns the matched node.
        Document ownerDoc = (Document)args.MatchNode.Document;

        // Create a DocumentBuilder positioned at the start of the matched node.
        DocumentBuilder builder = new DocumentBuilder(ownerDoc);
        builder.MoveTo(args.MatchNode);

        // Insert a hyperlink with the email as display text and a mailto: link.
        builder.InsertHyperlink(email, $"mailto:{email}", false);

        // Remove the original run that contained the plain email text.
        Node originalNode = args.MatchNode;
        originalNode.Remove();

        // Increment the replacement counter.
        Replacements++;

        // Skip the default replacement because we have already inserted the hyperlink.
        return ReplaceAction.Skip;
    }
}
