using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json;

public class MacroExpander
{
    // Custom callback that expands macros using the provided dictionary.
    private class MacroReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _macroMap;

        public MacroReplacer(Dictionary<string, string> macroMap)
        {
            _macroMap = macroMap ?? throw new ArgumentNullException(nameof(macroMap));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The regex has a single capturing group that contains the macro name.
            string macroName = args.Match.Groups[1].Value;

            if (_macroMap.TryGetValue(macroName, out string replacement))
                args.Replacement = replacement;          // Replace with the mapped value.
            else
                args.Replacement = args.Match.Value;      // Leave unknown macros unchanged.

            return ReplaceAction.Replace;
        }
    }

    public static void Main()
    {
        // Create a blank document and add sample text containing macro placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated on {DATE}.");
        builder.Writeln("Prepared by {FULLNAME}.");
        builder.Writeln("Message: {GREETING}");
        builder.Writeln("Unknown macro: {UNKNOWN}");

        // Define macro names and their full code replacements.
        var macroMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DATE", DateTime.Now.ToString("yyyy-MM-dd") },
            { "FULLNAME", "John Doe" },
            { "GREETING", "Hello, world!" }
        };

        // Regular expression to locate macros in the form {MACRO_NAME}.
        Regex macroRegex = new Regex(@"\{(\w+)\}", RegexOptions.Compiled);

        // Set up FindReplaceOptions with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new MacroReplacer(macroMap));

        // Perform the replace. The replacement string argument is ignored because the callback supplies it.
        int replacedCount = doc.Range.Replace(macroRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No macros were expanded.");

        // Save the modified document.
        const string outputPath = "ExpandedMacros.docx";
        doc.Save(outputPath);

        // Optional: Serialize the macro map to JSON for demonstration purposes.
        string json = JsonConvert.SerializeObject(macroMap, Formatting.Indented);
        System.IO.File.WriteAllText("MacroMap.json", json);
    }
}
