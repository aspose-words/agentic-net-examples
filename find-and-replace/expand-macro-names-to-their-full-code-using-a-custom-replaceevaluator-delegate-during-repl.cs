using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with macro placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {{NAME}},");
        builder.Writeln("Your appointment is on {{DATE}}.");
        builder.Writeln("Please review the code: {{CODE_SNIPPET}}.");

        // Save the sample input document locally.
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define macro expansions.
        var macroMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "NAME", "John Doe" },
            { "DATE", DateTime.Today.ToString("D") },
            { "CODE_SNIPPET", "Console.WriteLine(\"Hello World\");" }
        };

        // Regular expression to locate macros like {{MACRO}}.
        Regex macroRegex = new Regex(@"\{\{(\w+)\}\}");

        // Set up FindReplaceOptions with a custom callback that expands macros.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new MacroReplacingCallback(macroMap, macroRegex);

        // Perform the replacement. The replacement string argument is ignored because the callback supplies the actual text.
        int replacedCount = loaded.Range.Replace(macroRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No macros were replaced.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Custom callback that replaces each macro with its corresponding value from the dictionary.
    private class MacroReplacingCallback : IReplacingCallback
    {
        private readonly Dictionary<string, string> _macroMap;
        private readonly Regex _regex;

        public MacroReplacingCallback(Dictionary<string, string> macroMap, Regex regex)
        {
            _macroMap = macroMap;
            _regex = regex;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match found by the regex.
            Match match = args.Match;
            string key = match.Groups[1].Value;

            // Look up the macro value; if not found, keep the original placeholder.
            if (_macroMap.TryGetValue(key, out string replacement))
            {
                args.Replacement = replacement;
            }
            else
            {
                args.Replacement = match.Value;
            }

            return ReplaceAction.Replace;
        }
    }
}
