using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing macros.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello ${NAME}, today is ${DATE}.");
        builder.Writeln("Your order number is ${ORDER_ID}.");
        const string inputPath = "input.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Regex to locate macros like ${MACRO_NAME}
        var macroRegex = new Regex(@"\$\{(\w+)\}", RegexOptions.Compiled);

        // Collect information about each replacement for reporting.
        var replacements = new List<ReplacementInfo>();

        // Set up the callback that will expand macros and record the operation.
        var callback = new MacroReplacingCallback(replacements);
        var options = new FindReplaceOptions { ReplacingCallback = callback };

        // Perform the replace operation using the callback.
        int replacedCount = loadedDoc.Range.Replace(macroRegex, string.Empty, options);

        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one macro replacement, but none occurred.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Serialize replacement report to JSON.
        const string reportPath = "replacements.json";
        var json = JsonConvert.SerializeObject(replacements, Formatting.Indented);
        File.WriteAllText(reportPath, json);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
        {
            throw new InvalidOperationException("Failed to create the replacements report file.");
        }

        // Example completed – no interactive prompts.
    }

    private class ReplacementInfo
    {
        public string Macro { get; set; } = string.Empty;
        public string Replacement { get; set; } = string.Empty;
    }

    private class MacroReplacingCallback : IReplacingCallback
    {
        private readonly List<ReplacementInfo> _replacements;

        public MacroReplacingCallback(List<ReplacementInfo> replacements)
        {
            _replacements = replacements ?? throw new ArgumentNullException(nameof(replacements));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract macro name from the match.
            var match = args.Match;
            var macroName = match.Groups[1].Value;

            // Determine replacement text.
            string replacement = macroName switch
            {
                "NAME" => "John Doe",
                "DATE" => DateTime.Now.ToString("yyyy-MM-dd"),
                "ORDER_ID" => "12345",
                _ => match.Value // Unknown macro – leave unchanged.
            };

            // Record the replacement.
            _replacements.Add(new ReplacementInfo
            {
                Macro = macroName,
                Replacement = replacement
            });

            // Apply the replacement.
            args.Replacement = replacement;
            return ReplaceAction.Replace;
        }
    }
}
