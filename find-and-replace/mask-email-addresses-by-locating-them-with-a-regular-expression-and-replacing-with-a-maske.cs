using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current directory.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "input.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");

        // -----------------------------------------------------------------
        // Create a sample document containing e‑mail addresses.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please contact john.doe@example.com or jane_smith@domain.org for assistance.");
        builder.Writeln("Another address: admin@my-company.co.uk");
        doc.Save(inputPath); // Save the source document.

        // -----------------------------------------------------------------
        // Load the document we just created.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regular expression that matches typical e‑mail addresses.
        Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", RegexOptions.Compiled);

        // Set up find‑replace options with a custom callback that masks the e‑mail.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EmailMaskCallback()
        };

        // Perform the replacement. The replacement string argument is ignored because the callback supplies the value.
        int replacedCount = loaded.Range.Replace(emailRegex, string.Empty, options);

        // Verify that at least one e‑mail address was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No e‑mail addresses were found to mask.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Simple console output to indicate success.
        Console.WriteLine($"Masked {replacedCount} e‑mail address(es). Output saved to: {outputPath}");
    }

    // Callback that replaces each matched e‑mail with a masked version.
    private class EmailMaskCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Original e‑mail address.
            string original = args.Match.Value;

            // Split into local part and domain.
            int atIndex = original.IndexOf('@');
            if (atIndex <= 0)
                return ReplaceAction.Skip; // Guard against malformed matches.

            string localPart = original.Substring(0, atIndex);
            string domainPart = original.Substring(atIndex + 1);

            // Mask the local part with asterisks, preserving its length.
            string maskedLocal = new string('*', localPart.Length);
            string maskedEmail = $"{maskedLocal}@{domainPart}";

            // Set the replacement text.
            args.Replacement = maskedEmail;
            return ReplaceAction.Replace;
        }
    }
}
