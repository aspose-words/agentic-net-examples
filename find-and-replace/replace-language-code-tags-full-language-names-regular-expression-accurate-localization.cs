using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Regex to find language code tags in the form {1234}
        Regex languageTagPattern = new Regex(@"\{(\d+)\}");

        // Set up replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new LanguageTagReplacer()
        };

        // Perform the replace operation. The replacement string is ignored because the callback supplies it.
        doc.Range.Replace(languageTagPattern, string.Empty, options);

        // Save the updated document.
        doc.Save("Output.docx");
    }
}

// Callback that converts a numeric language code into its full language name.
class LanguageTagReplacer : IReplacingCallback
{
    // Cache of language code (int) to enum name for quick lookup.
    private static readonly Dictionary<int, string> _codeToName = BuildCodeToNameMap();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Extract the numeric code captured by the regex.
        string codeStr = args.Match.Groups[1].Value;
        if (int.TryParse(codeStr, out int code) && _codeToName.TryGetValue(code, out string languageName))
        {
            // Replace the whole tag with the language name.
            args.Replacement = languageName;
        }
        else
        {
            // If the code is unknown, keep the original tag.
            args.Replacement = args.Match.Value;
        }

        return ReplaceAction.Replace;
    }

    // Builds a dictionary mapping enum values (language codes) to their enum names.
    private static Dictionary<int, string> BuildCodeToNameMap()
    {
        var map = new Dictionary<int, string>();
        foreach (var value in Enum.GetValues(typeof(Aspose.Words.AI.Language)))
        {
            int intValue = (int)value;
            string name = Enum.GetName(typeof(Aspose.Words.AI.Language), value);
            if (!map.ContainsKey(intValue))
                map.Add(intValue, name);
        }
        return map;
    }
}
