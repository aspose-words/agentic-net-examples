using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some text containing macro placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated on DATE_PLACEHOLDER.");
        builder.Writeln("Please review the following sections:");
        builder.Writeln("MACRO_SUMMARY");
        builder.Writeln("MACRO_DETAILS");
        builder.Writeln("End of report.");

        // Define a regular expression that matches the macro placeholders.
        // In this example macros are simple words in uppercase without spaces.
        Regex macroPattern = new Regex(@"\b(MACRO_[A-Z]+|DATE_PLACEHOLDER)\b");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new MacroExpander()
        };

        // Perform the replacement.
        int replacementCount = doc.Range.Replace(macroPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No macro placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "ExpandedMacros.docx";
        doc.Save(outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Custom callback that expands macro names to their full code.
    private class MacroExpander : IReplacingCallback
    {
        // Mapping of macro placeholders to their expanded text.
        private readonly Dictionary<string, string> _macroMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DATE_PLACEHOLDER", DateTime.Now.ToString("D") },
            { "MACRO_SUMMARY", "Summary:\n- Item A\n- Item B\n- Item C" },
            { "MACRO_DETAILS", "Details:\n1. Detail one.\n2. Detail two.\n3. Detail three." }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string macroName = args.Match.Value;
            if (_macroMap.TryGetValue(macroName, out string replacement))
            {
                args.Replacement = replacement;
                return ReplaceAction.Replace;
            }

            // If the macro is unknown, leave it unchanged.
            return ReplaceAction.Skip;
        }
    }
}
