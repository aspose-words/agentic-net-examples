using System;
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
        builder.Writeln("This is a sample document.");
        builder.Writeln("Current date macro: [DATE]");
        builder.Writeln("User name macro: [USERNAME]");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up a callback that expands macros to their full values.
        var macroCallback = new MacroExpander();
        FindReplaceOptions options = new FindReplaceOptions(macroCallback);

        // Find macros of the form [MACRO_NAME] using a regular expression.
        Regex macroPattern = new Regex(@"\[([A-Z]+)\]");

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(macroPattern, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No macro placeholders were found for replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that replaces each macro with its expanded value.
    private class MacroExpander : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the macro name without brackets (captured group 1).
            string macroName = args.Match.Groups[1].Value;
            args.Replacement = ExpandMacro(macroName);
            return ReplaceAction.Replace;
        }

        private string ExpandMacro(string name) => name switch
        {
            "DATE" => DateTime.Now.ToString("yyyy-MM-dd"),
            "USERNAME" => Environment.UserName,
            _ => $"[UNKNOWN:{name}]"
        };
    }
}
