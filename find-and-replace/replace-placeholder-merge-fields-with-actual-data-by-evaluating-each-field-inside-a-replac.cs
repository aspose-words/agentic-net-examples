using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class FindAndReplaceDemo
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeFindReplaceDemo");
        Directory.CreateDirectory(workFolder);

        // Paths for the sample input, output and report files.
        string inputPath = Path.Combine(workFolder, "input.docx");
        string outputPath = Path.Combine(workFolder, "output.docx");
        string reportPath = Path.Combine(workFolder, "report.json");

        // Create a sample document containing placeholder merge fields.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear {{FirstName}} {{LastName}},");
        builder.Writeln("Your order {{OrderId}} is confirmed.");
        builder.Writeln("Thank you for shopping with us.");
        sampleDoc.Save(inputPath);

        // Load the document to be processed.
        Document doc = new Document(inputPath);

        // Data that will replace the placeholders.
        var placeholderValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "OrderId", "12345" }
        };

        // Keep track of which placeholders were actually replaced.
        var replacedPlaceholders = new List<string>();

        // Define a regex that matches {{Placeholder}} patterns.
        Regex placeholderRegex = new Regex(@"{{(\w+)}}", RegexOptions.Compiled);

        // Set up the callback that performs the replacement.
        var callback = new PlaceholderReplacer(placeholderValues, replacedPlaceholders);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement. The replacement string argument is ignored when a callback is used.
        int replaceCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replaceCount == 0 || replacedPlaceholders.Count == 0)
        {
            throw new InvalidOperationException("No placeholders were replaced. Expected at least one replacement.");
        }

        // Save the modified document.
        doc.Save(outputPath);

        // Create a simple JSON report of the performed replacements.
        var report = new
        {
            ReplacedPlaceholders = replacedPlaceholders,
            ReplacementCount = replaceCount,
            OutputDocument = outputPath
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(reportPath, jsonReport);

        // Verify that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new InvalidOperationException("Failed to create the replacement report.");
        }

        // The example runs to completion without requiring any user interaction.
    }

    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _values;
        private readonly IList<string> _replaced;

        public PlaceholderReplacer(IDictionary<string, string> values, IList<string> replaced)
        {
            _values = values ?? throw new ArgumentNullException(nameof(values));
            _replaced = replaced ?? throw new ArgumentNullException(nameof(replaced));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the placeholder name without the braces.
            string key = args.Match.Groups[1].Value;
            if (_values.TryGetValue(key, out string replacement))
            {
                args.Replacement = replacement;
                _replaced.Add(key);
                return ReplaceAction.Replace;
            }

            // No replacement found – keep the original text.
            return ReplaceAction.Skip;
        }
    }
}
