using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare sample data for placeholder replacement.
        var data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FullName", "John Doe" },
            { "Date", DateTime.Today.ToString("yyyy-MM-dd") },
            { "Company", "Acme Corp" }
        };

        // Create a new document with placeholder merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear <<FullName>>,");
        builder.Writeln("Welcome to <<Company>>.");
        builder.Writeln("Your appointment is scheduled for <<Date>>.");
        builder.Writeln("Thank you.");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new PlaceholderReplacer(data));

        // Use a regular expression to locate placeholders of the form <<FieldName>>.
        Regex placeholderPattern = new Regex(@"<<\w+>>", RegexOptions.Compiled);

        // Perform the replacement.
        int replacements = doc.Range.Replace(placeholderPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacements == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the resulting document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Replaced {replacements} placeholders. Document saved to '{outputPath}'.");
    }

    // Implements IReplacingCallback to evaluate each placeholder and provide the replacement text.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Extract the placeholder name without the surrounding << and >>.
            string placeholder = args.Match.Value; // e.g., "<<FullName>>"
            string key = placeholder.Substring(2, placeholder.Length - 4); // "FullName"

            // Look up the replacement value; if not found, keep the original placeholder.
            if (_values.TryGetValue(key, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = placeholder; // fallback to original text

            return ReplaceAction.Replace;
        }
    }
}
