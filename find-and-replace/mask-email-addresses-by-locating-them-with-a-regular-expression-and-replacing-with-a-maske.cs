using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with email addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact list:");
        builder.Writeln("Alice: alice@example.com");
        builder.Writeln("Bob: bob.smith@subdomain.example.org");
        builder.Writeln("Charlie: charlie123@example.co.uk");
        builder.Writeln("No email here.");

        // Define a regular expression that matches typical email addresses.
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", RegexOptions.Compiled);

        // Set up find‑replace options with a custom callback that masks the email.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EmailMaskingCallback()
        };

        // Perform the replacement. The replacement string is ignored when a callback is used.
        int replacedCount = doc.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one email was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found to mask.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MaskedEmails.docx");
        doc.Save(outputPath);
    }

    // Callback that replaces each found email with a masked version.
    private class EmailMaskingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The original email address.
            string originalEmail = args.Match.Value;

            // Split into local part and domain.
            int atIndex = originalEmail.IndexOf('@');
            if (atIndex <= 0)
            {
                // If the format is unexpected, leave it unchanged.
                args.Replacement = originalEmail;
                return ReplaceAction.Replace;
            }

            string localPart = originalEmail.Substring(0, atIndex);
            string domainPart = originalEmail.Substring(atIndex + 1);

            // Mask the local part with asterisks, preserving its length.
            string maskedLocal = new string('*', localPart.Length);

            // Construct the masked email.
            string maskedEmail = $"{maskedLocal}@{domainPart}";

            // Set the replacement text.
            args.Replacement = maskedEmail;
            return ReplaceAction.Replace;
        }
    }
}
