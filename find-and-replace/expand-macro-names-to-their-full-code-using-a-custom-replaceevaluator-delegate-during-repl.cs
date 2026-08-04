using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with macro placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {{FirstName}} {{LastName}}!");
        builder.Writeln("Today is {{Day}}.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define macro expansions.
        var macroMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "Alice" },
            { "LastName", "Smith" },
            { "Day", "Monday" }
        };

        // Set up find/replace with a custom callback.
        var options = new FindReplaceOptions(new MacroExpander(macroMap));
        int replacedCount = loadedDoc.Range.Replace(new Regex(@"\{\{(\w+)\}\}"), string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one macro expansion.");

        // Save the expanded document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Write a simple JSON report of performed expansions.
        var report = new
        {
            InputFile = Path.GetFullPath(inputPath),
            OutputFile = Path.GetFullPath(outputPath),
            Replacements = macroMap,
            ReplacedCount = replacedCount
        };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }

    // Custom callback that replaces macro placeholders with their full code.
    private class MacroExpander : IReplacingCallback
    {
        private readonly IDictionary<string, string> _macroMap;
        public MacroExpander(IDictionary<string, string> macroMap) => _macroMap = macroMap;

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The regex captures the macro name in group 1.
            string macroName = args.Match.Groups[1].Value;
            if (_macroMap.TryGetValue(macroName, out string replacement))
                args.Replacement = replacement;
            else
                args.Replacement = args.Match.Value; // Leave unchanged if not found.

            return ReplaceAction.Replace;
        }
    }
}
