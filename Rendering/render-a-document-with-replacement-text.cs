using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsReplacementDemo
{
    // Custom callback that replaces placeholders with values from a dictionary
    // and logs each replacement.
    public class PlaceholderReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _values;
        private readonly StringBuilder _log = new StringBuilder();

        public PlaceholderReplacer(Dictionary<string, string> values)
        {
            _values = values;
        }

        // This method is called for each match found by the Find/Replace engine.
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The whole match includes the surrounding brackets, e.g. "[NAME]".
            string placeholder = args.Match.Value.Trim('[', ']');

            // Determine the replacement text.
            if (_values.TryGetValue(placeholder, out string replacement))
            {
                args.Replacement = replacement;
                _log.AppendLine($"Replaced \"{args.Match.Value}\" with \"{replacement}\".");
                return ReplaceAction.Replace;
            }

            // If no replacement is found, skip this match.
            _log.AppendLine($"No replacement found for \"{args.Match.Value}\". Skipping.");
            return ReplaceAction.Skip;
        }

        public string GetLog() => _log.ToString();
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Insert sample text containing placeholders.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Dear [NAME],");
            builder.Writeln("Welcome to [PLACE]!");
            builder.Writeln("Your appointment is on [DATE].");

            // Define the values that will replace the placeholders.
            var replacements = new Dictionary<string, string>
            {
                { "NAME", "John Doe" },
                { "PLACE", "Aspose City" },
                { "DATE", DateTime.Today.ToString("MMMM d, yyyy") }
            };

            // Set up find/replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new PlaceholderReplacer(replacements)
            };

            // Perform the replacement using a regular expression that matches [PLACEHOLDER] patterns.
            doc.Range.Replace(new Regex(@"\[\w+\]"), string.Empty, options);

            // Save the resulting document.
            doc.Save("ReplacedDocument.docx");
        }
    }
}
