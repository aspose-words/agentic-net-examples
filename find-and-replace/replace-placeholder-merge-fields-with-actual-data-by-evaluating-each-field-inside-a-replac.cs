using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    // Callback that replaces {{FieldName}} placeholders with values from a dictionary.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The regular expression has one capturing group that contains the field name.
            string key = args.Match.Groups[1].Value;

            if (_values.TryGetValue(key, out string replacement))
            {
                // Set the replacement text and tell Aspose.Words to replace the match.
                args.Replacement = replacement;
                return ReplaceAction.Replace;
            }

            // No matching key – leave the original placeholder unchanged.
            return ReplaceAction.Skip;
        }
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing placeholder merge fields.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear {{FirstName}} {{LastName}},");
        builder.Writeln("Your order {{OrderId}} has been shipped.");
        builder.Writeln("Thank you for shopping with us.");

        // Save the source document (optional, just to illustrate the input file).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Define the data that will replace the placeholders.
        // -----------------------------------------------------------------
        var placeholderValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "OrderId", "12345" }
        };

        // -----------------------------------------------------------------
        // 3. Set up a regular expression that matches {{FieldName}} patterns.
        // -----------------------------------------------------------------
        Regex placeholderRegex = new Regex(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

        // -----------------------------------------------------------------
        // 4. Perform the replacement using a custom IReplacingCallback.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new PlaceholderReplacer(placeholderValues)
        };

        // The replacement string argument is ignored because the callback sets args.Replacement.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
