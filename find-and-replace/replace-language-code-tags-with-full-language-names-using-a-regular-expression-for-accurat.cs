using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class LanguageTagReplacer : IReplacingCallback
{
    private readonly Dictionary<string, string> _languageMap;

    public LanguageTagReplacer(Dictionary<string, string> languageMap)
    {
        _languageMap = languageMap ?? throw new ArgumentNullException(nameof(languageMap));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The regex has one capturing group that contains the language code.
        string code = args.Match.Groups[1].Value;
        if (_languageMap.TryGetValue(code, out string fullName))
        {
            args.Replacement = fullName;
            return ReplaceAction.Replace;
        }

        // If the code is not in the map, leave the original text unchanged.
        return ReplaceAction.Skip;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document containing language code tags like [en], [fr], etc.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome to our site. [en]");
        builder.Writeln("Bienvenue sur notre site. [fr]");
        builder.Writeln("Bienvenido a nuestro sitio. [es]");
        builder.Writeln("Willkommen auf unserer Seite. [de]");

        // Define the mapping from language codes to full language names.
        var languageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en", "English" },
            { "fr", "French" },
            { "es", "Spanish" },
            { "de", "German" }
        };

        // Regular expression to find tags like [en], [fr], etc.
        Regex tagRegex = new Regex(@"\[(\w{2})\]", RegexOptions.Compiled);

        // Set up find‑replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new LanguageTagReplacer(languageMap);

        // Perform the replacement. The replacement string argument is ignored because the callback supplies it.
        int replacedCount = doc.Range.Replace(tagRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {Path.GetFullPath(outputPath)}");
    }
}
