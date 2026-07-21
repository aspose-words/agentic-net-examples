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
        builder.Writeln("Please contact john.doe@example.com or jane_smith@domain.org for assistance.");
        builder.Writeln("Another address: admin@my-company.net.");

        // Define a regular expression that matches e‑mail addresses.
        const string emailPattern = @"\b[\w\.-]+@[\w\.-]+\.\w+\b";

        // Set up find‑replace options with a custom callback that masks each e‑mail.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EmailMaskCallback()
        };

        // Perform the replacement. The replacement string is ignored because the callback supplies the masked value.
        int replacedCount = doc.Range.Replace(new Regex(emailPattern), string.Empty, options);

        // Ensure that at least one e‑mail address was found and masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No e‑mail addresses were found to mask.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that receives each e‑mail match and replaces it with a masked version.
    private class EmailMaskCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string email = args.Match.Value;
            int atIndex = email.IndexOf('@');
            if (atIndex > 1)
            {
                // Keep the first character of the local part, mask the rest, and keep the domain unchanged.
                string localPart = email.Substring(0, atIndex);
                string domainPart = email.Substring(atIndex + 1);
                string maskedLocal = localPart[0] + new string('*', localPart.Length - 1);
                args.Replacement = $"{maskedLocal}@{domainPart}";
            }
            else
            {
                // Fallback masking if the e‑mail format is unexpected.
                args.Replacement = "*****@*****";
            }

            return ReplaceAction.Replace;
        }
    }
}
