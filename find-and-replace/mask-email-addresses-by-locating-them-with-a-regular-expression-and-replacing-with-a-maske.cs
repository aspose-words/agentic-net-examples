using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.Collections.Generic;
using Newtonsoft.Json;

public class EmailMaskCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Get the matched email address.
        string email = args.Match.Value;
        int atIndex = email.IndexOf('@');
        if (atIndex > 0)
        {
            // Mask the local part with asterisks.
            string maskedLocal = new string('*', atIndex);
            string domain = email.Substring(atIndex);
            args.Replacement = maskedLocal + domain;
        }
        else
        {
            // If somehow the pattern matched without '@', leave it unchanged.
            args.Replacement = email;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document containing email addresses.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please contact support@example.com for assistance.");
        builder.Writeln("You can also reach sales@mydomain.org or info@company.co.uk.");
        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches email addresses.
        Regex emailRegex = new Regex(@"\b[\w\.-]+@[\w\.-]+\.\w{2,}\b", RegexOptions.Compiled);

        // Set up find-and-replace options with a custom callback to mask emails.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EmailMaskCallback()
        };

        // Perform the replacement. The replacement string parameter is ignored because the callback sets it.
        int replacedCount = loaded.Range.Replace(emailRegex, string.Empty, options);

        // Validate that at least one email address was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No email addresses were found to mask.");

        // Save the resulting document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
