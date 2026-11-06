using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json;

public class MacroReplacingCallback : IReplacingCallback
{
    private readonly Dictionary<string, string> _macros;

    public MacroReplacingCallback(Dictionary<string, string> macros)
    {
        _macros = macros ?? throw new ArgumentNullException(nameof(macros));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The matched macro text, e.g. {{DATE}}
        string key = args.Match.Value;

        // Look up the replacement value; if not found, keep the original macro.
        if (_macros.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            args.Replacement = key;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing macro placeholders.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated on {{DATE}} by {{USER}}.");
        builder.Writeln("Contact: {{EMAIL}}.");
        const string inputPath = "input.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Define macro expansions.
        // -----------------------------------------------------------------
        var macroValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "{{DATE}}", DateTime.Now.ToString("yyyy-MM-dd") },
            { "{{USER}}", Environment.UserName },
            { "{{EMAIL}}", $"{Environment.UserName}@example.com" }
        };

        // -----------------------------------------------------------------
        // 4. Set up regex to locate macros and the callback that expands them.
        // -----------------------------------------------------------------
        Regex macroRegex = new Regex(@"{{\w+}}", RegexOptions.Compiled);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new MacroReplacingCallback(macroValues)
        };

        // -----------------------------------------------------------------
        // 5. Perform the replacement using the callback.
        // -----------------------------------------------------------------
        // The replacement string is ignored because the callback supplies the actual text.
        int replacedCount = loaded.Range.Replace(macroRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No macros were expanded; expected at least one replacement.");

        // -----------------------------------------------------------------
        // 6. Save the modified document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        loaded.Save(outputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 7. Write a JSON report of the performed replacements.
        // -----------------------------------------------------------------
        var report = new
        {
            InputFile = Path.GetFullPath(inputPath),
            OutputFile = Path.GetFullPath(outputPath),
            ReplacementsMade = replacedCount,
            ExpandedMacros = macroValues
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("replacement-report.json", jsonReport);
    }
}
