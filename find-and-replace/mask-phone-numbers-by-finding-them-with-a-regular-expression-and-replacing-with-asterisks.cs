using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing phone numbers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Contact list:");
        builder.Writeln("John Doe: 123-456-7890");
        builder.Writeln("Jane Smith: (555) 123 4567");
        builder.Writeln("International: +1 800 555 0199");

        // Define a regular expression that matches common phone number formats.
        string phonePattern = @"\b(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{3}\)?[-.\s]?)?\d{3}[-.\s]?\d{4}\b";

        // Set up find-and-replace options with a custom callback that masks each match.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PhoneMaskCallback();

        // Perform the replacement. The replacement string is ignored because the callback provides its own.
        int replacementCount = doc.Range.Replace(new Regex(phonePattern), string.Empty, options);

        // Validate that at least one phone number was masked.
        if (replacementCount == 0)
            throw new InvalidOperationException("No phone numbers were found to mask.");

        // Save the modified document.
        const string outputPath = "MaskedPhoneNumbers.docx";
        doc.Save(outputPath);
    }

    // Callback that replaces each phone number with a string of asterisks of the same length.
    private class PhoneMaskCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            args.Replacement = new string('*', args.Match.Value.Length);
            return ReplaceAction.Replace;
        }
    }
}
