using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for font handling

public class Program
{
    public static void Main()
    {
        // Create a sample document with language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome messages:");
        builder.Writeln("<en> Hello!");
        builder.Writeln("<fr> Bonjour!");
        builder.Writeln("<es> Hola!");
        builder.Writeln("<de> Guten Tag!");
        // Save the original for reference (optional).
        doc.Save("input.docx");

        // Define a callback that replaces each language code with its full name.
        LanguageTagReplacer replacer = new LanguageTagReplacer();

        // Set up find‑replace options to use the callback.
        FindReplaceOptions options = new FindReplaceOptions(replacer);

        // Regular expression to match tags like <en>, <fr>, etc.
        Regex tagRegex = new Regex(@"<([a-z]{2})>", RegexOptions.IgnoreCase);

        // Perform the replacement. The replacement string is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(tagRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        doc.Save("output.docx");

        // Output the result count (console output is allowed for logging).
        Console.WriteLine($"Replaced {replacedCount} language tag(s).");
    }

    // Callback that maps language codes to full language names.
    private class LanguageTagReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> LanguageMap = new()
        {
            { "en", "English" },
            { "fr", "French" },
            { "es", "Spanish" },
            { "de", "German" },
            { "it", "Italian" },
            { "pt", "Portuguese" },
            { "ru", "Russian" },
            { "zh", "Chinese" },
            { "ja", "Japanese" },
            { "ko", "Korean" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the language code from the first capture group.
            string code = args.Match.Groups[1].Value.ToLowerInvariant();

            // Look up the full language name; if not found, keep the original tag.
            if (LanguageMap.TryGetValue(code, out string fullName))
            {
                args.Replacement = fullName;
                return ReplaceAction.Replace;
            }

            // No mapping found – skip replacement.
            return ReplaceAction.Skip;
        }
    }
}
