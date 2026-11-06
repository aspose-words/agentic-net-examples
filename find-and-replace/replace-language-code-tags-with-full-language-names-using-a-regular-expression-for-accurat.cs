using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ReplaceLanguageCallback : IReplacingCallback
{
    private readonly Dictionary<string, string> _languageMap = new()
    {
        { "en", "English" },
        { "fr", "French" },
        { "es", "Spanish" },
        { "de", "German" },
        { "it", "Italian" }
        // Add more mappings as needed
    };

    public int ReplacementsMade { get; private set; }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract the language code from the match (e.g., "en" from "[en]")
        string code = args.Match.Groups[1].Value;

        if (_languageMap.TryGetValue(code, out string fullName))
        {
            // Set the replacement text to the full language name
            args.Replacement = fullName;
            ReplacementsMade++;
        }
        else
        {
            // If the code is unknown, keep the original tag
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document containing language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome message: [en]");
        builder.Writeln("Message en français: [fr]");
        builder.Writeln("Mensaje en español: [es]");
        builder.Writeln("Unbekannte Sprache: [xx]"); // This tag will remain unchanged.

        // Define a regular expression to match tags like [en], [fr], etc.
        Regex languageTagRegex = new Regex(@"\[([a-z]{2})\]", RegexOptions.IgnoreCase);

        // Set up the callback that will replace each tag with the full language name.
        ReplaceLanguageCallback callback = new ReplaceLanguageCallback();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement using the regex and callback.
        int totalReplacements = doc.Range.Replace(languageTagRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replacements performed: {callback.ReplacementsMade}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
