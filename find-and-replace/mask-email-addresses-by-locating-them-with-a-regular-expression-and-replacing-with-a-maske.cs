using System;
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
        builder.Writeln("Contact us at support@example.com or sales@example.org.");
        builder.Writeln("Personal email: john.doe123@mail.example.co.uk");

        // Define a regular expression that matches email addresses.
        Regex emailRegex = new Regex(@"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b");

        // Set up find‑replace options with a custom callback that masks the email.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new EmailMaskingCallback();

        // Perform the replacement. The replacement string is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one email was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found to mask.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that replaces each matched email with a masked version.
    private class EmailMaskingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string email = args.Match.Value;
            int atIndex = email.IndexOf('@');
            if (atIndex > 0)
            {
                // Mask the local part of the email, keep the domain unchanged.
                string maskedLocal = new string('*', atIndex);
                string maskedEmail = maskedLocal + email.Substring(atIndex);
                args.Replacement = maskedEmail;
            }
            else
            {
                // Fallback: replace the whole match with asterisks if the format is unexpected.
                args.Replacement = new string('*', email.Length);
            }

            return ReplaceAction.Replace;
        }
    }
}
