using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class LanguageTagReplacer : IReplacingCallback
{
    private readonly Dictionary<string, string> _languageMap = new()
    {
        { "en", "English" },
        { "fr", "French" },
        { "es", "Spanish" },
        { "de", "German" },
        { "it", "Italian" },
        { "pt", "Portuguese" }
    };

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // args.Match.Value includes the brackets, e.g., "[en]"
        string code = args.Match.Value.Trim('[', ']');
        if (_languageMap.TryGetValue(code, out string fullName))
        {
            args.Replacement = fullName;
        }
        else
        {
            // If the code is unknown, keep it unchanged.
            args.Replacement = args.Match.Value;
        }
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome [en] user!");
        builder.Writeln("Bonjour [fr] utilisateur!");
        builder.Writeln("¡Hola [es] usuario!");
        builder.Writeln("Hallo [de] Benutzer!");
        builder.Writeln("Ciao [it] utente!");
        builder.Writeln("Olá [pt] usuário!");
        builder.Writeln("Unknown [xx] tag stays.");

        // Save the original for reference (optional).
        doc.Save("original.docx");

        // Prepare regex to find tags like [en], [fr], etc.
        Regex tagRegex = new Regex(@"\[(\w{2})\]");

        // Set up find-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LanguageTagReplacer()
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(tagRegex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        doc.Save("localized.docx");

        // Output a simple verification to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine("Resulting text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
