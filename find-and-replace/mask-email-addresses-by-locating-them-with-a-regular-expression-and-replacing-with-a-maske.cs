using System;
using System.IO;
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
        builder.Writeln("Contact us at support@example.com for assistance.");
        builder.Writeln("Personal email: john.doe@mydomain.org");
        builder.Writeln("Another address: admin@company.co.uk");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Define a regular expression that matches typical email addresses.
        Regex emailRegex = new Regex(@"\b[\w.%+-]+@[\w.-]+\.[A-Za-z]{2,}\b");

        // Set up find‑replace options with a custom callback that masks each email.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new EmailMaskingCallback();

        // Perform the replacement. The replacement string is ignored because the callback supplies it.
        int replacedCount = loaded.Range.Replace(emailRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found to mask.");

        // Save the masked document.
        loaded.Save("output.docx");
    }

    // Callback that replaces each found email with a masked version.
    private class EmailMaskingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The full email address matched by the regex.
            string email = args.Match.Value;

            // Split the email into local part and domain.
            int atIndex = email.IndexOf('@');
            if (atIndex <= 0)
            {
                // If the format is unexpected, leave it unchanged.
                args.Replacement = email;
                return ReplaceAction.Replace;
            }

            string localPart = email.Substring(0, atIndex);
            string domainPart = email.Substring(atIndex); // includes '@'

            // Mask the local part with asterisks of the same length.
            string maskedLocal = new string('*', localPart.Length);
            args.Replacement = maskedLocal + domainPart;

            return ReplaceAction.Replace;
        }
    }
}
