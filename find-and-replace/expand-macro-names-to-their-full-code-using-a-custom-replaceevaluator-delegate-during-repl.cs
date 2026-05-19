using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing macro placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated on {DATE} by {USER}.");
        builder.Writeln("Contact: {EMAIL}");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loadedDoc = new Document(inputPath);

        // Define macro expansions.
        var macroValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DATE", DateTime.Now.ToString("yyyy-MM-dd") },
            { "USER", Environment.UserName },
            { "EMAIL", $"{Environment.UserName}@example.com" }
        };

        // Regular expression to find macros like {MACRO_NAME}.
        Regex macroRegex = new Regex(@"\{(\w+)\}", RegexOptions.Compiled);

        // Set up FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new MacroReplacingCallback(macroValues);

        // Perform the replacement using the regex pattern.
        int replacedCount = loadedDoc.Range.Replace(macroRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No macros were expanded in the document.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Confirmation output.
        Console.WriteLine($"Macros expanded: {replacedCount}");
        Console.WriteLine($"Output saved to '{outputPath}'.");
    }

    // Custom callback that expands macros based on the provided dictionary.
    private class MacroReplacingCallback : IReplacingCallback
    {
        private readonly Dictionary<string, string> _macroValues;

        public MacroReplacingCallback(Dictionary<string, string> macroValues)
        {
            _macroValues = macroValues ?? throw new ArgumentNullException(nameof(macroValues));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the macro name from the first capture group.
            string key = args.Match.Groups[1].Value;

            // If the macro exists, replace it with its value; otherwise keep the original text.
            if (_macroValues.TryGetValue(key, out string value))
                args.Replacement = value;
            else
                args.Replacement = args.Match.Value;

            return ReplaceAction.Replace;
        }
    }
}
