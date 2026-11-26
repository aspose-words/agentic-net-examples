using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class PlaceholderReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public PlaceholderReplacer(IDictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract the field name without the surrounding << and >>.
        string placeholder = args.Match.Value;
        string key = placeholder.Substring(2, placeholder.Length - 4);

        if (_values.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            // Keep the original placeholder if no replacement is found.
            args.Replacement = placeholder;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------------------
        // 1. Create a sample template document containing placeholder fields.
        // -------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Dear <<FirstName>> <<LastName>>,");
        builder.Writeln("Your order <<OrderNumber>> has been shipped.");
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // -------------------------------------------------------------
        // 2. Load the template for processing.
        // -------------------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------------------
        // 3. Define the data that will replace the placeholders.
        // -------------------------------------------------------------
        var fieldValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "OrderNumber", "12345" }
        };

        // -------------------------------------------------------------
        // 4. Prepare a regular expression that matches placeholders like <<FieldName>>.
        // -------------------------------------------------------------
        var placeholderRegex = new Regex(@"<<\w+>>", RegexOptions.Compiled);

        // -------------------------------------------------------------
        // 5. Set up the replacing callback that supplies the actual replacement text.
        // -------------------------------------------------------------
        var replacer = new PlaceholderReplacer(fieldValues);
        var options = new FindReplaceOptions { ReplacingCallback = replacer };

        // -------------------------------------------------------------
        // 6. Perform the replacement using the callback overload.
        // -------------------------------------------------------------
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // -------------------------------------------------------------
        // 7. Validate that at least one replacement occurred.
        // -------------------------------------------------------------
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // -------------------------------------------------------------
        // 8. Save the resulting document.
        // -------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // -------------------------------------------------------------
        // 9. Write a simple JSON report of the performed replacements.
        // -------------------------------------------------------------
        var report = new
        {
            InputFile = Path.GetFullPath(templatePath),
            OutputFile = Path.GetFullPath(outputPath),
            ReplacementsMade = replacedCount,
            ReplacedFields = fieldValues
        };
        File.WriteAllText("report.json", JsonConvert.SerializeObject(report, Formatting.Indented));
    }
}
