using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing placeholder tokens.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Your order {{OrderId}} has been shipped.");
        builder.Writeln("Thank you, {{Company}}.");

        // Save the original document for inspection (optional).
        doc.Save("input.docx");

        // Define the placeholder values to substitute.
        var placeholders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Name", "John Doe" },
            { "OrderId", "12345" },
            { "Company", "Acme Corp" }
        };

        // Set up find‑replace options with a custom callback that performs the substitution.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer(placeholders);

        // Regular expression that matches tokens of the form {{Key}}.
        Regex placeholderRegex = new Regex(@"\{\{.+?\}\}", RegexOptions.Compiled);

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Ensure that at least one placeholder was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces each matched placeholder with the corresponding value from the dictionary.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the placeholder key without the surrounding braces, e.g. "{{Name}}" -> "Name".
            string match = args.Match.Value;
            string key = match.Trim('{', '}');

            // Look up the replacement value.
            if (_values.TryGetValue(key, out string replacement))
            {
                // Provide the replacement text to Aspose.Words.
                args.Replacement = replacement;
                return ReplaceAction.Replace;
            }

            // If no replacement is found, leave the original text unchanged.
            return ReplaceAction.Skip;
        }
    }
}
