using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Languages used in the document:");
        builder.Writeln("- English: [en]");
        builder.Writeln("- French: [fr]");
        builder.Writeln("- German: [de]");
        builder.Writeln("- Spanish: [es]");

        // Save the original document (optional, for inspection).
        doc.Save("input.docx");

        // Set up the replace callback that maps language codes to full names.
        FindReplaceOptions options = new FindReplaceOptions(new LanguageTagReplacer());

        // Regex to find tags like [en], [fr], etc.
        Regex tagPattern = new Regex(@"\[(\w{2})\]");

        // Perform the replacement. The replacement string is ignored because the callback provides the actual text.
        int replacedCount = doc.Range.Replace(tagPattern, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }
}

// Callback that replaces language code tags with their full language names.
public class LanguageTagReplacer : IReplacingCallback
{
    private static readonly Dictionary<string, string> LanguageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "en", "English" },
        { "fr", "French" },
        { "de", "German" },
        { "es", "Spanish" }
    };

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract the language code without brackets.
        string code = args.Match.Value.Trim('[', ']');

        // Look up the full language name; if not found, keep the original code.
        if (LanguageMap.TryGetValue(code, out string fullName))
            args.Replacement = fullName;
        else
            args.Replacement = code;

        return ReplaceAction.Replace;
    }
}
