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
        // Create a sample JSON configuration file with placeholder values.
        const string jsonPath = "config.json";
        string jsonContent = @"{
            ""FirstName"": ""John"",
            ""LastName"": ""Doe"",
            ""Company"": ""Acme Corp""
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // Load the JSON into a dictionary for quick lookup.
        var placeholderValues = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(jsonPath))
                                 ?? new Dictionary<string, string>();

        // Create a sample Word document containing placeholders surrounded by double brackets.
        const string inputPath = "input.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello [[FirstName]] [[LastName]], welcome to [[Company]].");
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Define a regex that matches placeholders like [[Key]].
        var placeholderRegex = new Regex(@"\[\[(.+?)\]\]");

        // Set up find‑replace options with a custom callback.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new PlaceholderReplacer(placeholderValues)
        };

        // Perform the replacement. The replacement string argument is ignored because the callback supplies it.
        int replacedCount = loadedDoc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced. Check the input document and JSON configuration.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

// Callback that replaces each matched placeholder with the corresponding value from the JSON dictionary.
public class PlaceholderReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public PlaceholderReplacer(IDictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // args.Match.Value includes the surrounding brackets, e.g., [[FirstName]].
        string placeholder = args.Match.Value;

        // Extract the key without the brackets.
        string key = placeholder.Substring(2, placeholder.Length - 4);

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
