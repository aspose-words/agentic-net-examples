using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsReplaceExample
{
    // Implements a custom callback that replaces tokens like {{Name}} using a dictionary.
    public class ReplaceEvaluator : IReplacingCallback
    {
        private readonly Dictionary<string, string> _values;

        public ReplaceEvaluator(Dictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The regex pattern captures the token name without the surrounding braces.
            // args.Match.Groups[1] contains the token (e.g., Name, Company).
            string token = args.Match.Groups[1].Value;

            if (_values.TryGetValue(token, out string replacement))
            {
                args.Replacement = replacement;
                return ReplaceAction.Replace;
            }

            // If the token is not found in the dictionary, leave it unchanged.
            return ReplaceAction.Skip;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample document containing placeholder tokens.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Hello {{Name}}!");
            builder.Writeln("Welcome to {{Company}}.");
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the document that we just created.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the replacement values.
            // -----------------------------------------------------------------
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Name", "John Doe" },
                { "Company", "Acme Corp" }
            };

            // -----------------------------------------------------------------
            // 4. Set up FindReplaceOptions with the custom callback.
            // -----------------------------------------------------------------
            var options = new FindReplaceOptions(new ReplaceEvaluator(replacements));

            // The regex matches tokens of the form {{TokenName}} and captures the name.
            Regex tokenRegex = new Regex(@"\{\{(\w+)\}\}");

            // Perform the replacement.
            int replacedCount = doc.Range.Replace(tokenRegex, string.Empty, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No placeholders were replaced.");

            // -----------------------------------------------------------------
            // 5. Save the modified document.
            // -----------------------------------------------------------------
            const string resultPath = "result.docx";
            doc.Save(resultPath);

            // Optional: output a simple confirmation.
            Console.WriteLine($"Replacements performed: {replacedCount}");
            Console.WriteLine($"Result saved to: {Path.GetFullPath(resultPath)}");
        }
    }
}
