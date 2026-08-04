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
        // Create a sample document with language code tags like [en], [fr], [es].
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This is an English sentence: [en]");
        builder.Writeln("Bonjour le monde! Ceci est une phrase française: [fr]");
        builder.Writeln("¡Hola mundo! Esta es una frase en español: [es]");
        builder.Writeln("Unsupported tag example: [xx]");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define a regular expression that matches language tags in the form [xx].
        Regex tagRegex = new Regex(@"\[(\w{2})\]", RegexOptions.Compiled);

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new LanguageTagReplacer());

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace(tagRegex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Optional: display the resulting text in the console.
        Console.WriteLine("Replacement completed. Resulting document text:");
        Console.WriteLine(loadedDoc.GetText().Trim());
    }

    // Callback that replaces a language code tag with its full language name.
    private class LanguageTagReplacer : IReplacingCallback
    {
        // Mapping from two‑letter language codes to full language names.
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
            // The regex has a capturing group for the language code.
            string code = args.Match.Groups[1].Value.ToLowerInvariant();

            // Look up the full language name; if not found, keep the original tag.
            if (LanguageMap.TryGetValue(code, out string fullName))
                args.Replacement = fullName;
            else
                args.Replacement = args.Match.Value; // leave unchanged

            return ReplaceAction.Replace;
        }
    }
}
