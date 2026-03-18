using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Dictionary that maps macro names to their full VBA source code.
        var macroCodeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "MyMacro", "Sub MyMacro()\n    MsgBox \"Hello\"\nEnd Sub" },
            { "ViewZoom200", "Sub ViewZoom200()\n    ActiveWindow.View.Zoom.Percentage = 200\nEnd Sub" }
        };

        // Create FindReplaceOptions and attach the custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new MacroExpander(macroCodeMap));

        // Regex that finds macro placeholders like {{MacroName}}.
        Regex macroPlaceholder = new Regex(@"\{\{(?<name>\w+)\}\}");

        // Perform the replace. The callback supplies the actual replacement text.
        doc.Range.Replace(macroPlaceholder, string.Empty, options);

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Custom callback that expands a macro name to its full VBA code.
    private class MacroExpander : IReplacingCallback
    {
        private readonly IDictionary<string, string> _macroMap;

        public MacroExpander(IDictionary<string, string> macroMap)
        {
            _macroMap = macroMap;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the macro name from the named capture group "name".
            string macroName = args.Match.Groups["name"].Value;

            // If the macro is known, replace the placeholder with its code.
            // Word paragraph breaks are represented by the meta‑character "&p".
            if (_macroMap.TryGetValue(macroName, out string code))
                args.Replacement = code.Replace("\n", "&p");
            else
                args.Replacement = args.Match.Value; // leave unchanged if unknown.

            return ReplaceAction.Replace;
        }
    }
}
