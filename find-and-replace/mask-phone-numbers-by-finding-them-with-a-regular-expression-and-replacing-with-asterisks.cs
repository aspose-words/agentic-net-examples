using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with phone numbers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact: John Doe, phone: 123-456-7890.");
        builder.Writeln("Support: +1 (800) 555-1234, alternate: 555.987.6543.");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Regular expression to match common phone number formats.
        Regex phoneRegex = new Regex(@"\+?\d{0,2}[\s\-\.\(]*\d{3}[\s\-\.\)]*\d{3}[\s\-\.\)]*\d{4}");

        // Callback that replaces each matched phone number with asterisks of equal length.
        IReplacingCallback maskCallback = new PhoneMaskCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = maskCallback
        };

        // Perform the replacement. The replacement string is ignored because the callback sets it.
        int replacedCount = loaded.Range.Replace(phoneRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No phone numbers were found to mask.");

        // Save the masked document.
        loaded.Save("output.docx");
    }

    // Callback implementation for masking phone numbers.
    private class PhoneMaskCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace each character of the matched phone number with '*'.
            args.Replacement = new string('*', args.Match.Value.Length);
            return ReplaceAction.Replace;
        }
    }
}
