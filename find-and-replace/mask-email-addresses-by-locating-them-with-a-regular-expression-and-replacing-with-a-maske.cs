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
        builder.Writeln("Contact us at support@example.com or sales@example.org.");
        builder.Writeln("Another email: user.name+tag@sub.domain.co.uk");
        builder.Writeln("No email on this line.");

        // Regular expression that matches typical email addresses.
        Regex emailRegex = new Regex(@"\b[\w\.\-+]+@[\w\.\-]+\.\w+\b", RegexOptions.Compiled);

        // Configure find‑replace to use a custom callback that masks the local part.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EmailMaskCallback()
        };

        // Perform the replacement. The replacement string is ignored when a callback is supplied.
        int replacedCount = doc.Range.Replace(emailRegex, string.Empty, options);

        // Verify that at least one email address was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one email address to be masked.");

        // Save the resulting document.
        const string outputPath = "masked-emails.docx";
        doc.Save(outputPath);
    }

    // Callback that replaces the local part of an e‑mail address with asterisks.
    private class EmailMaskCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Guard against a null match.
            if (args?.Match == null)
                return ReplaceAction.Skip;

            string email = args.Match.Value;
            int atIndex = email.IndexOf('@');
            if (atIndex <= 0)
                return ReplaceAction.Skip; // Not a valid e‑mail format.

            // Create a masked local part of the same length.
            string maskedLocal = new string('*', atIndex);
            string maskedEmail = maskedLocal + email.Substring(atIndex);

            // Set the replacement text.
            args.Replacement = maskedEmail;
            return ReplaceAction.Replace;
        }
    }
}
