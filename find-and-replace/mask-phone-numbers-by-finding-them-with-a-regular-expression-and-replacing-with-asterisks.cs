using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various phone number formats.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact list:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: (987) 654-3210");
        builder.Writeln("Bob Johnson: 5551234567");
        builder.Writeln("Alice Brown: 321.654.0987");

        // Define a regular expression that matches common US phone number patterns.
        Regex phoneRegex = new Regex(@"\b(?:\d{3}[-.\s]?\d{3}[-.\s]?\d{4}|\(\d{3}\)\s?\d{3}[-.\s]?\d{4})\b");

        // Set up find‑replace options with a custom callback that masks the matched phone number.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PhoneMaskCallback();

        // Perform the replacement. The replacement string argument is ignored because the callback provides the value.
        int replacedCount = doc.Range.Replace(phoneRegex, string.Empty, options);

        // Validate that at least one phone number was masked.
        if (replacedCount == 0)
            throw new InvalidOperationException("No phone numbers were found to mask.");

        // Save the modified document.
        doc.Save("masked.docx");
    }

    // Callback that replaces each matched phone number with a string of asterisks of the same length.
    private class PhoneMaskCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string original = args.Match.Value;
            string masked = new string('*', original.Length);
            args.Replacement = masked;
            return ReplaceAction.Replace;
        }
    }
}
