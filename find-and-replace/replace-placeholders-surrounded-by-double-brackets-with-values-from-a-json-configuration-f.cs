using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public class Program
{
    public static void Main()
    {
        // Prepare a JSON configuration file with placeholder values.
        const string jsonPath = "config.json";
        var jsonContent = @"{
            ""FirstName"": ""John"",
            ""LastName"": ""Doe"",
            ""Date"": ""2024-12-31""
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // Load the JSON into a dictionary for quick lookup.
        var placeholderValues = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(jsonPath))
                                 ?? new Dictionary<string, string>();

        // Create a sample Word document containing placeholders surrounded by double brackets.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello [[FirstName]] [[LastName]],");
        builder.Writeln("today is [[Date]].");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform find-and-replace.
        var loadedDoc = new Document(inputPath);

        // Set up the callback that will replace each placeholder with the corresponding JSON value.
        var replacer = new PlaceholderReplacer(placeholderValues);
        var options = new FindReplaceOptions
        {
            ReplacingCallback = replacer
        };

        // Regex matches any text like [[Placeholder]] (non‑greedy).
        var placeholderPattern = new Regex(@"\[\[(.+?)\]\]");

        // Perform the replacement. The actual replacement text is supplied by the callback.
        int replacedCount = loadedDoc.Range.Replace(placeholderPattern, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Optional: display the resulting document text to verify the operation.
        Console.WriteLine("Replaced document text:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }
}

// Callback that replaces matched placeholders with values from a dictionary.
public class PlaceholderReplacer : IReplacingCallback
{
    private readonly Dictionary<string, string> _values;

    public PlaceholderReplacer(Dictionary<string, string> values)
    {
        _values = values ?? new Dictionary<string, string>();
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // args.Match.Value includes the surrounding brackets, e.g., [[FirstName]].
        string match = args.Match.Value;
        // Extract the key between the brackets.
        string key = match.Length > 4 ? match.Substring(2, match.Length - 4) : string.Empty;

        if (_values.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            // If the key is not found, keep the original placeholder.
            args.Replacement = match;
        }

        return ReplaceAction.Replace;
    }
}
