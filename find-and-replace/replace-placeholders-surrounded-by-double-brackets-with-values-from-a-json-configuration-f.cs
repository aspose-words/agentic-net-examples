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
        // Prepare a simple JSON configuration file.
        const string jsonPath = "config.json";
        var jsonContent = @"{
            ""FirstName"": ""John"",
            ""LastName"": ""Doe"",
            ""OrderId"": ""12345""
        }";
        File.WriteAllText(jsonPath, jsonContent);

        // Load configuration into a dictionary.
        var config = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonContent)
                     ?? new Dictionary<string, string>();

        // Create a sample document containing placeholders surrounded by double brackets.
        const string inputPath = "input.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello [[FirstName]] [[LastName]]!");
        builder.Writeln("Your order [[OrderId]] is confirmed.");
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Set up find‑replace options with a custom callback that substitutes placeholders.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new JsonPlaceholderReplacer(config)
        };

        // Use a regular expression to locate placeholders like [[Key]].
        var placeholderPattern = new Regex(@"\[\[(.+?)\]\]");
        int replacedCount = loadedDoc.Range.Replace(placeholderPattern, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

// Callback that replaces each matched placeholder with the corresponding value from the JSON configuration.
public class JsonPlaceholderReplacer : IReplacingCallback
{
    private readonly IReadOnlyDictionary<string, string> _values;

    public JsonPlaceholderReplacer(IReadOnlyDictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // The first capturing group contains the key without the surrounding brackets.
        var key = args.Match.Groups[1].Value;
        if (_values.TryGetValue(key, out var replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            // If the key is not found, keep the original placeholder.
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }
}
