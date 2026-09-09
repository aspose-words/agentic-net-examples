using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with numeric values in US format (e.g., 1,234.56).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample numbers:");
        builder.Writeln("1,234.56");
        builder.Writeln("12,345.00");
        builder.Writeln("987.65");
        builder.Writeln("No number here.");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Regular expression to match numbers with optional thousands separators and a decimal part.
        // Example matches: 1,234.56 , 12,345 , 987.65
        Regex numberRegex = new Regex(@"\b\d{1,3}(?:,\d{3})*(?:\.\d+)?\b");

        // Callback that converts the matched US‑style number to a German‑style format (e.g., 1.234,56).
        NumberLocalizationCallback callback = new NumberLocalizationCallback(new CultureInfo("de-DE"));

        // Configure find‑replace options to use the callback.
        FindReplaceOptions options = new FindReplaceOptions(callback);

        // Perform the replacement. The callback supplies the actual replacement string.
        int replacedCount = loaded.Range.Replace(numberRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Write a JSON report of the original and replaced values.
        string jsonReport = JsonConvert.SerializeObject(callback.Replacements, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}

// Holds information about a single replacement operation.
public class ReplacementInfo
{
    public string Original { get; set; } = string.Empty;
    public string Replacement { get; set; } = string.Empty;
}

// Callback that parses the matched number using invariant culture and formats it using a target culture.
public class NumberLocalizationCallback : IReplacingCallback
{
    private readonly List<ReplacementInfo> _replacements = new List<ReplacementInfo>();
    private readonly CultureInfo _targetCulture;

    public NumberLocalizationCallback(CultureInfo targetCulture)
    {
        _targetCulture = targetCulture ?? throw new ArgumentNullException(nameof(targetCulture));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        string original = args.Match.Value;

        // Try to parse the number assuming US formatting (comma as thousands separator, dot as decimal separator).
        if (double.TryParse(original, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint,
                            CultureInfo.InvariantCulture, out double number))
        {
            // Format the number using the target culture (e.g., German uses '.' for thousands and ',' for decimals).
            string formatted = number.ToString("N", _targetCulture);
            args.Replacement = formatted;

            _replacements.Add(new ReplacementInfo
            {
                Original = original,
                Replacement = formatted
            });
        }
        else
        {
            // If parsing fails, keep the original text unchanged.
            args.Replacement = original;
        }

        return ReplaceAction.Replace;
    }

    // Expose the collected replacement data for reporting.
    public IReadOnlyList<ReplacementInfo> Replacements => _replacements;
}
