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
        // Create a sample document with placeholder merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear {FirstName} {LastName},");
        builder.Writeln("Welcome to {Company}.");
        builder.Writeln("Your employee ID is {EmployeeId}.");

        // Prepare data that will replace the placeholders.
        var data = new Dictionary<string, string>
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "Company", "Acme Corp" },
            { "EmployeeId", "12345" }
        };

        // Serialize the data to a JSON file (demonstrates required package usage).
        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(data, Formatting.Indented));

        // Define a regular expression that matches placeholders like {FieldName}.
        Regex placeholderRegex = new Regex(@"\{(\w+)\}");

        // Configure find/replace options and assign a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new PlaceholderReplacer(data)
        };

        // Perform the replace operation. The replacement string is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Simple verification that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Custom callback that replaces each placeholder with the corresponding value from the dictionary.
    private class PlaceholderReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _data;

        public PlaceholderReplacer(IDictionary<string, string> data)
        {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Extract the field name from the first capturing group.
            string fieldName = args.Match.Groups[1].Value;

            // Determine the replacement value.
            if (_data.TryGetValue(fieldName, out string value))
                args.Replacement = value;
            else
                args.Replacement = args.Match.Value; // Keep original placeholder if not found.

            return ReplaceAction.Replace;
        }
    }
}
