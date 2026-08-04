using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document containing placeholder tokens.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {{FirstName}} {{LastName}}! Welcome to {{Company}}.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document we just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Define the replacement values in a dictionary.
        // -----------------------------------------------------------------
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName",  "Doe" },
            { "Company",   "Acme Corp" }
        };

        // -----------------------------------------------------------------
        // 4. Set up a callback that replaces each token with the dictionary value.
        // -----------------------------------------------------------------
        var callback = new TokenReplaceCallback(values);
        var options = new FindReplaceOptions(callback);

        // Regex that matches tokens like {{TokenName}}.
        Regex tokenRegex = new Regex(@"\{\{[A-Za-z0-9_]+\}\}");

        // Perform the find-and-replace operation.
        int replacedCount = loadedDoc.Range.Replace(tokenRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one placeholder replacement.");

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // Optional: output the final document text to the console.
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }

    // -----------------------------------------------------------------
    // Callback implementation that looks up token values in a dictionary.
    // -----------------------------------------------------------------
    private class TokenReplaceCallback : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public TokenReplaceCallback(IDictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The matched token, e.g., "{{FirstName}}".
            string token = args.Match.Value;

            // Extract the key without the surrounding braces.
            string key = token.Trim('{', '}');

            // Look up the replacement value; if not found, keep the original token.
            if (_values.TryGetValue(key, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = token;

            return ReplaceAction.Replace;
        }
    }
}
