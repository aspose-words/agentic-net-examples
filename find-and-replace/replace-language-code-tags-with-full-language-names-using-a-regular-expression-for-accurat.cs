using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required for Aspose.Words formatting APIs

public class Program
{
    public static void Main()
    {
        // Create a sample document with language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome [en] user!");
        builder.Writeln("¡Bienvenido [es] usuario!");
        builder.Writeln("Bienvenue [fr] utilisateur!");
        builder.Writeln("Willkommen [de] Benutzer!");

        // Define a regex that matches tags like [en], [es], etc.
        Regex languageTagRegex = new Regex(@"\[([a-z]{2})\]", RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new LanguageTagReplacer();

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = doc.Range.Replace(languageTagRegex, string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Callback that converts a language code (e.g., "en") to its full name (e.g., "English").
    private class LanguageTagReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> LanguageMap = new()
        {
            { "en", "English" },
            { "es", "Spanish" },
            { "fr", "French" },
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
                args.Replacement = fullName;
            else
                args.Replacement = args.Match.Value; // fallback to original tag

            return ReplaceAction.Replace;
        }
    }
}
