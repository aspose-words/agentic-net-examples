using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json;

public class Program
{
    // Custom callback that replaces <<FieldName>> placeholders with values from a dictionary.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _data;

        public PlaceholderReplacer(Dictionary<string, string> data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The regex has one capturing group that contains the field name.
            string key = args.Match.Groups[1].Value;

            if (_data.TryGetValue(key, out string value))
                args.Replacement = value;          // Replace with the found value.
            else
                args.Replacement = args.Match.Value; // Keep the original placeholder if not found.

            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a new blank document and write sample text containing placeholder merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear <<FirstName>> <<LastName>>,");
        builder.Writeln("Your employee ID is <<EmployeeID>>.");
        builder.Writeln("Welcome to <<Company>>.");

        // Data that will replace the placeholders.
        var placeholderData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "EmployeeID", "12345" },
            { "Company", "Acme Corp" }
        };

        // Regular expression that matches placeholders of the form <<FieldName>>.
        Regex placeholderRegex = new Regex(@"<<(\w+)>>", RegexOptions.Compiled);

        // Set up FindReplaceOptions with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new PlaceholderReplacer(placeholderData)
        };

        // Perform the replace operation. The replacement string argument is ignored because
        // the callback supplies the actual replacement text.
        int replacementCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Create a simple JSON report of the operation.
        var report = new
        {
            ReplacementsMade = replacementCount,
            DataUsed = placeholderData
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);

        // Output information to the console (no interactive input required).
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
        Console.WriteLine($"Report saved to: {Path.GetFullPath("report.json")}");
    }
}
