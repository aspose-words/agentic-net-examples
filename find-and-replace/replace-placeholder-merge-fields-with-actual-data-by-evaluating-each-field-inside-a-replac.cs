using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string inputPath = Path.Combine(workDir, "input.docx");
        string outputPath = Path.Combine(workDir, "output.docx");
        string reportPath = Path.Combine(workDir, "report.json");

        // -----------------------------------------------------------------
        // 1. Create a sample document containing placeholder merge fields.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear {{FirstName}} {{LastName}},");
        builder.Writeln("Your order {{OrderId}} has been shipped.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document that we just created.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Define the data that will replace the placeholders.
        // -----------------------------------------------------------------
        var data = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName",  "Doe" },
            { "OrderId",   "A12345" }
        };

        // -----------------------------------------------------------------
        // 4. Perform a find‑and‑replace using a custom IReplacingCallback.
        //    The regular expression matches {{FieldName}} patterns.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer(data);

        // The replacement string is ignored because the callback supplies the actual text.
        int replacedCount = loaded.Range.Replace(new Regex(@"{{(\w+)}}"), string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // -----------------------------------------------------------------
        // 6. Write a simple JSON report of the performed replacements.
        // -----------------------------------------------------------------
        var report = new
        {
            InputFile = inputPath,
            OutputFile = outputPath,
            ReplacementsMade = replacedCount,
            ReplacedFields = data
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(reportPath, jsonReport);
    }

    // Custom callback that replaces {{FieldName}} with values from a dictionary.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;

        public PlaceholderReplacer(IDictionary<string, string> values)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The first captured group contains the field name without the braces.
            string key = args.Match.Groups[1].Value;

            if (_values.TryGetValue(key, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = args.Match.Value; // keep original if not found.

            return ReplaceAction.Replace;
        }
    }
}
