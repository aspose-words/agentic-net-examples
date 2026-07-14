using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplaceEvaluator : IReplacingCallback
{
    private readonly Dictionary<string, string> _replacements;

    public ReplaceEvaluator(Dictionary<string, string> replacements)
    {
        _replacements = replacements ?? new Dictionary<string, string>();
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The regex captures the token name without the surrounding braces.
        // Group 1 contains the key (e.g., Name, Company).
        string token = args.Match.Groups[1].Value;

        if (_replacements.TryGetValue(token, out string value))
        {
            args.Replacement = value;
        }
        else
        {
            // If the token is not found, keep the original placeholder.
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";
        const string jsonPath = "data.json";

        // 1. Create a sample document containing placeholder tokens.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {{Name}}, welcome to {{Company}}.");
        doc.Save(inputPath);

        // 2. Load the document from the file system.
        Document loaded = new Document(inputPath);

        // 3. Prepare the replacement values.
        var replacements = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Company", "Acme Corp" }
        };

        // Optional: serialize the dictionary to JSON for reporting purposes.
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(replacements, Formatting.Indented));

        // 4. Set up the find-and-replace operation using a regex and a callback evaluator.
        var evaluator = new ReplaceEvaluator(replacements);
        var options = new FindReplaceOptions(evaluator);
        var regex = new Regex(@"\{\{(\w+)\}\}");

        int replacedCount = loaded.Range.Replace(regex, string.Empty, options);

        // 5. Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // 6. Save the modified document.
        loaded.Save(outputPath);
    }
}
