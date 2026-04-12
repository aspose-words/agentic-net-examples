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
        // Prepare a temporary folder for all files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a JSON configuration file with placeholder values.
        string jsonPath = Path.Combine(workDir, "config.json");
        var configData = new Dictionary<string, string>
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "Company", "Aspose Ltd." }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(configData, Formatting.Indented));

        // Load the JSON into a dictionary for quick lookup.
        var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(jsonPath))
                     ?? new Dictionary<string, string>();

        // Create a sample Word document containing placeholders surrounded by double brackets.
        string templatePath = Path.Combine(workDir, "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Dear [[FirstName]] [[LastName]],");
        builder.Writeln("Welcome to [[Company]]!");
        doc.Save(templatePath);

        // Load the document to perform find‑and‑replace.
        var loadedDoc = new Document(templatePath);

        // Set up find‑replace options with a custom callback.
        var options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer(values);

        // Regex that matches [[PlaceholderName]].
        var pattern = new Regex(@"\[\[(\w+)\]\]");

        // Perform the replace operation.
        int replaceCount = loadedDoc.Range.Replace(pattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replaceCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the resulting document.
        string resultPath = Path.Combine(workDir, "Result.docx");
        loadedDoc.Save(resultPath);

        // Optional: indicate success (no interactive console required).
        Console.WriteLine($"Replacements performed: {replaceCount}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }

    // Callback that replaces each matched placeholder with the corresponding value from the JSON.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values)
        {
            _values = values ?? new Dictionary<string, string>();
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Extract the placeholder name from the first capturing group.
            string key = args.Match.Groups[1].Value;

            if (_values.TryGetValue(key, out string replacement))
            {
                args.Replacement = replacement;
                return ReplaceAction.Replace;
            }

            // If the key is not found, skip this match to leave the original text unchanged.
            return ReplaceAction.Skip;
        }
    }
}
