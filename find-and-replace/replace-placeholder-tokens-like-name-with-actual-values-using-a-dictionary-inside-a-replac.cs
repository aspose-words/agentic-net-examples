using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly in this example.

namespace FindAndReplaceExample
{
    // Implements a callback that replaces placeholders with values from a dictionary.
    public class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match will be something like "{{Name}}".
            string placeholder = args.Match.Value;

            // Extract the key without the surrounding braces.
            // Assumes the placeholder format is exactly {{Key}}.
            string key = placeholder.Length > 4
                ? placeholder.Substring(2, placeholder.Length - 4)
                : string.Empty;

            // Look up the replacement value; if not found, keep the original placeholder.
            if (_values.TryGetValue(key, out string replacement))
            {
                args.Replacement = replacement;
            }
            else
            {
                args.Replacement = placeholder;
            }

            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a sample document containing placeholder tokens.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Dear {{Name}},");
            builder.Writeln("Your order number {{OrderId}} has been shipped on {{Date}}.");
            builder.Writeln("Thank you for shopping with us!");
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Step 2: Load the document we just created.
            Document loaded = new Document(inputPath);

            // Step 3: Prepare the replacement values.
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Name", "John Doe" },
                { "OrderId", "12345" },
                { "Date", DateTime.Today.ToString("d") }
            };

            // Step 4: Set up the callback and find‑replace options.
            var replacer = new PlaceholderReplacer(values);
            var options = new FindReplaceOptions
            {
                ReplacingCallback = replacer
            };

            // Step 5: Define a regex that matches placeholders of the form {{Key}}.
            Regex placeholderRegex = new Regex(@"{{\w+}}");

            // Perform the replacement. The replacement string argument is ignored because the callback supplies the actual text.
            int replacedCount = loaded.Range.Replace(placeholderRegex, string.Empty, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No placeholders were replaced. Expected at least one replacement.");

            // Step 6: Save the modified document.
            const string outputPath = "output.docx";
            loaded.Save(outputPath);

            // Optional: Write a simple confirmation to the console.
            Console.WriteLine($"Replacements performed: {replacedCount}");
            Console.WriteLine($"Modified document saved to '{outputPath}'.");
        }
    }
}
