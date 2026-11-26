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
        // 1. Create a sample template document containing placeholders.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello [[Name]],");
        builder.Writeln("Your order [[OrderId]] has been shipped on [[ShipDate]].");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a JSON configuration file with replacement values.
        // -----------------------------------------------------------------
        const string jsonPath = "config.json";
        var config = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "OrderId", "12345" },
            { "ShipDate", DateTime.Today.ToString("yyyy-MM-dd") }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(config));

        // -----------------------------------------------------------------
        // 3. Load the JSON configuration into a dictionary.
        // -----------------------------------------------------------------
        string jsonContent = File.ReadAllText(jsonPath);
        var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonContent)
                     ?? throw new InvalidOperationException("Failed to deserialize JSON configuration.");

        // -----------------------------------------------------------------
        // 4. Load the template document that contains the placeholders.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Set up the callback that will replace each placeholder with the
        //    corresponding value from the JSON configuration.
        // -----------------------------------------------------------------
        var replacer = new PlaceholderReplacer(values);
        var options = new FindReplaceOptions { ReplacingCallback = replacer };

        // -----------------------------------------------------------------
        // 6. Perform the replacement using a regular expression that matches
        //    the pattern [[Key]].
        // -----------------------------------------------------------------
        Regex placeholderRegex = new Regex(@"\[\[(.+?)\]\]");
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // -----------------------------------------------------------------
        // 7. Validate that at least one replacement occurred.
        // -----------------------------------------------------------------
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // -----------------------------------------------------------------
        // 8. Save the resulting document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 9. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}

// ---------------------------------------------------------------------
// Callback implementation that replaces placeholders with values
// from the JSON configuration.
// ---------------------------------------------------------------------
public class PlaceholderReplacer : IReplacingCallback
{
    private readonly IReadOnlyDictionary<string, string> _values;

    public PlaceholderReplacer(IReadOnlyDictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The match value includes the surrounding brackets, e.g., [[Name]].
        string placeholder = args.Match.Value;

        // Extract the key between the brackets.
        string key = placeholder.Length > 4 ? placeholder.Substring(2, placeholder.Length - 4) : string.Empty;

        if (string.IsNullOrEmpty(key) || !_values.TryGetValue(key, out string replacement))
        {
            // No matching key; skip this match.
            return ReplaceAction.Skip;
        }

        // Replace the matched text with the replacement string.
        // Aspose.Words will create a new Run node for the replacement text.
        args.Replacement = replacement;

        // Preserve original formatting if the match node is a Run.
        if (args.MatchNode is Run originalRun)
        {
            // The replacement string will be inserted as a new Run; we can
            // copy the formatting from the original run to the new run after
            // the replacement is performed by handling the NodeInserted event,
            // but for simplicity we rely on Aspose.Words to keep the formatting
            // of the original run when possible.
        }

        return ReplaceAction.Replace;
    }
}
