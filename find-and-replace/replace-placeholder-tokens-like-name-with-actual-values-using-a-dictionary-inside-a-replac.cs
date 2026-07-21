using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words drawing dependencies
using Newtonsoft.Json; // Included as per required packages

namespace AsposeWordsReplaceExample
{
    // Implements a custom callback that replaces placeholders with values from a dictionary.
    public class PlaceholderReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _values;

        public PlaceholderReplacer(Dictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The match will be something like "{{Name}}".
            string placeholder = args.Match.Value;

            // Extract the key without the surrounding braces.
            string key = placeholder.Length > 4
                ? placeholder.Substring(2, placeholder.Length - 4) // Remove leading "{{" and trailing "}}"
                : string.Empty;

            // Look up the key in the dictionary; if not found, keep the original placeholder.
            if (_values.TryGetValue(key, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = placeholder;

            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a sample document containing placeholder tokens.
            const string templatePath = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Dear {{Name}},");
            builder.Writeln("Welcome to {{Company}}.");
            builder.Writeln("Your role: {{Title}}.");
            templateDoc.Save(templatePath);

            // Step 2: Define the placeholder values.
            var placeholderValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Name", "John Doe" },
                { "Company", "Acme Corporation" },
                { "Title", "Senior Engineer" }
            };

            // Step 3: Load the document and perform the replace operation using a callback.
            Document doc = new Document(templatePath);
            var replaceOptions = new FindReplaceOptions
            {
                ReplacingCallback = new PlaceholderReplacer(placeholderValues),
                MatchCase = false,
                FindWholeWordsOnly = false
            };

            // Regex matches any token of the form {{Word}}.
            var placeholderRegex = new Regex(@"\{\{(\w+)\}\}");
            int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, replaceOptions);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No placeholders were replaced. Check the template and dictionary.");

            // Step 4: Save the resulting document.
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            // Optional: Write a simple confirmation to the console.
            Console.WriteLine($"Replaced {replacedCount} placeholder(s). Output saved to '{outputPath}'.");
        }
    }
}
