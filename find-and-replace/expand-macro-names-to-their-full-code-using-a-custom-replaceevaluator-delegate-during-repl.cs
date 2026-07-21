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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write sample text containing macro placeholders.
        builder.Writeln("The following macro will be expanded:");
        builder.Writeln("[[HELLO_WORLD]]");
        builder.Writeln("Another macro:");
        builder.Writeln("[[CURRENT_DATE]]");

        // Define macro expansions.
        var macroMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "HELLO_WORLD", "Console.WriteLine(\"Hello, World!\");" },
            { "CURRENT_DATE", "Console.WriteLine(DateTime.Now.ToString(\"yyyy-MM-dd\"));" }
        };

        // Set up find-and-replace with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new MacroExpander(macroMap);

        // Use a regular expression to locate macro placeholders like [[MACRO_NAME]].
        Regex macroRegex = new Regex(@"\[\[(\w+)\]\]");

        int replacedCount = doc.Range.Replace(macroRegex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No macro placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "ExpandedMacros.docx";
        doc.Save(outputPath);

        // Output the resulting document text to the console for verification.
        Console.WriteLine("Replacements performed: " + replacedCount);
        Console.WriteLine("Document text after replacement:");
        Console.WriteLine(doc.GetText());
    }

    // Implements IReplacingCallback to replace each macro with its full code.
    private class MacroExpander : IReplacingCallback
    {
        private readonly IDictionary<string, string> _macroMap;

        public MacroExpander(IDictionary<string, string> macroMap)
        {
            _macroMap = macroMap ?? throw new ArgumentNullException(nameof(macroMap));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // args.Match.Value contains the full placeholder, e.g., [[HELLO_WORLD]].
            // Extract the macro name without brackets.
            string placeholder = args.Match.Value;
            string macroName = placeholder.Trim('[', ']');

            if (_macroMap.TryGetValue(macroName, out string replacement))
            {
                // Set the replacement text.
                args.Replacement = replacement;
            }
            else
            {
                // If macro not found, keep original text.
                args.Replacement = placeholder;
            }

            return ReplaceAction.Replace;
        }
    }
}
