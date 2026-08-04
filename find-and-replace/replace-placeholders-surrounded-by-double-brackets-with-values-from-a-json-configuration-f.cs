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
        // 1. Create a JSON configuration file with placeholder values.
        // -----------------------------------------------------------------
        const string jsonConfig = @"{
            ""FirstName"": ""John"",
            ""LastName"": ""Doe"",
            ""Date"": ""2023-12-31""
        }";

        const string configPath = "config.json";
        File.WriteAllText(configPath, jsonConfig);

        // Parse the JSON into a dictionary.
        IDictionary<string, string> values = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(configPath))
                                            ?? new Dictionary<string, string>();

        // -----------------------------------------------------------------
        // 2. Create a sample Word document containing placeholders.
        // -----------------------------------------------------------------
        const string inputPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello [[FirstName]] [[LastName]], today is [[Date]].");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 3. Load the document and replace placeholders using a callback.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);
        var replacer = new PlaceholderReplacer(values);
        var options = new FindReplaceOptions { ReplacingCallback = replacer };

        // Regex matches any text surrounded by double brackets, e.g. [[Key]].
        int replacedCount = loaded.Range.Replace(new Regex(@"\[\[(.+?)\]\]"), string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one placeholder replacement.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: display the resulting document text.
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}

// ---------------------------------------------------------------------
// Callback that replaces each matched placeholder with the value from
// the JSON configuration. If a key is missing, the original placeholder
// is left unchanged.
// ---------------------------------------------------------------------
public class PlaceholderReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public PlaceholderReplacer(IDictionary<string, string> values)
    {
        _values = values ?? new Dictionary<string, string>();
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // args.Match.Value includes the surrounding brackets, e.g. [[FirstName]].
        string placeholder = args.Match.Value;

        // Extract the key between the brackets.
        // Length is at least 4 (e.g. [[a]]).
        string key = placeholder.Substring(2, placeholder.Length - 4);

        if (_values.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            // Keep the original placeholder if no matching key is found.
            args.Replacement = placeholder;
        }

        return ReplaceAction.Replace;
    }
}
