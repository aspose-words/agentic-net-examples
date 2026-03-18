using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the Word template that contains placeholders like [[FirstName]]
        Document doc = new Document("Template.docx");

        // Read the JSON configuration file and deserialize it into a dictionary
        string json = File.ReadAllText("config.json");
        var values = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

        // Prepare find/replace options and assign the custom callback
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new JsonPlaceholderReplacer(values);

        // Regular expression to match placeholders surrounded by double brackets
        Regex placeholderPattern = new Regex(@"\[\[(.+?)\]\]");

        // Perform the replacement across the whole document
        doc.Range.Replace(placeholderPattern, string.Empty, options);

        // Save the resulting document
        doc.Save("Result.docx");
    }
}

// Custom callback that substitutes each matched placeholder with the corresponding JSON value
class JsonPlaceholderReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public JsonPlaceholderReplacer(IDictionary<string, string> values)
    {
        _values = values;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Extract the key without the surrounding brackets (e.g., "FirstName" from "[[FirstName]]")
        string key = args.Match.Groups[1].Value;

        // Look up the key in the JSON dictionary; if not found, keep the original placeholder
        if (_values.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }
}
