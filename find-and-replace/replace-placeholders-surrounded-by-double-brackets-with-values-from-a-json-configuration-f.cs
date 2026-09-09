using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing placeholders like [[Name]]
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello [[FirstName]] [[LastName]]! Your order [[OrderId]] is confirmed.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // ---------------------------------------------------------------
        // 2. Load replacement values from a JSON configuration (in‑memory)
        // ---------------------------------------------------------------
        const string json = @"{ ""FirstName"": ""John"", ""LastName"": ""Doe"", ""OrderId"": ""12345"" }";
        Dictionary<string, string> values = JsonConvert.DeserializeObject<Dictionary<string, string>>(json)
                                          ?? new Dictionary<string, string>();

        // ---------------------------------------------------------------
        // 3. Load the document and replace placeholders using a callback
        // ---------------------------------------------------------------
        Document loaded = new Document(inputPath);
        FindReplaceOptions options = new FindReplaceOptions(new PlaceholderReplacer(values));

        // Regex matches [[Placeholder]] and captures the name inside the brackets
        int replacedCount = loaded.Range.Replace(new Regex(@"\[\[(.+?)\]\]"), string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // ---------------------------------------------------------------
        // 4. Save the modified document
        // ---------------------------------------------------------------
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // -----------------------------------------------------------------
    // Callback that substitutes each matched placeholder with the value
    // from the JSON dictionary.
    // -----------------------------------------------------------------
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values) => _values = values;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // args.Match.Value is the whole match, e.g. [[FirstName]]
            // Group 1 contains the placeholder name without brackets.
            string placeholderName = args.Match.Groups[1].Value;

            if (_values.TryGetValue(placeholderName, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = args.Match.Value; // keep original if not found

            return ReplaceAction.Replace;
        }
    }
}
