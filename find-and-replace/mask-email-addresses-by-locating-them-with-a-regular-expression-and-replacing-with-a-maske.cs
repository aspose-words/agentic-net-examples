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
        builder.Writeln("Contact us at support@example.com or sales@example.org.");
        builder.Writeln("Personal email: john.doe@mydomain.com.");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Regular expression to find email addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.Compiled);

        // Set up find‑replace options with a custom callback to mask the local part.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EmailMaskCallback()
        };

        // Perform the replacement. The replacement string is ignored when a callback is used.
        int replacedCount = loaded.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one email was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found to mask.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Optional: indicate completion (no interactive input).
        Console.WriteLine($"Masked {replacedCount} email address(es). Output saved to 'output.docx'.");
    }
}

// Callback that replaces each email with a masked version (e.g., *****@domain.com).
public class EmailMaskCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        string email = args.Match.Value;
        int atPos = email.IndexOf('@');
        if (atPos > 0)
        {
            string domain = email.Substring(atPos);
            string maskedLocal = new string('*', atPos);
            args.Replacement = maskedLocal + domain;
        }
        else
        {
            // Fallback: no '@' found, leave the original text unchanged.
            args.Replacement = email;
        }

        return ReplaceAction.Replace;
    }
}
