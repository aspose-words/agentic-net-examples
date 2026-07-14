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
        // Prepare file paths in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");

        // 1. Create a sample JSON configuration file.
        var config = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Date", DateTime.Today.ToString("yyyy-MM-dd") },
            { "Company", "Acme Corp" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(config));

        // 2. Create a sample Word document containing placeholders like [[Name]].
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello [[Name]],");
        builder.Writeln("Welcome to [[Company]] on [[Date]].");
        doc.Save(inputPath);

        // 3. Load the JSON configuration into a dictionary.
        var jsonText = File.ReadAllText(jsonPath);
        var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText)
                     ?? new Dictionary<string, string>();

        // 4. Load the document we just created.
        var loadedDoc = new Document(inputPath);

        // 5. Define a regex that matches placeholders surrounded by double brackets.
        var placeholderRegex = new Regex(@"\[\[(\w+)\]\]");

        // 6. Set up find‑replace options with a custom callback.
        var options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer(values);

        // 7. Perform the replace operation. The actual replacement text is supplied by the callback.
        int replacedCount = loadedDoc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // 8. Save the modified document.
        loadedDoc.Save(outputPath);
    }

    // Callback that replaces each matched placeholder with the corresponding value from the JSON.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _values;

        public PlaceholderReplacer(Dictionary<string, string> values) => _values = values;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The full match includes the brackets, e.g., [[Name]].
            // Group 1 captures the key without brackets.
            string key = args.Match.Groups[1].Value;

            if (_values.TryGetValue(key, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = args.Match.Value; // leave unchanged if key not found.

            return ReplaceAction.Replace;
        }
    }
}
