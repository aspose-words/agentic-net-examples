using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class AbbreviationReplacer : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The abbreviation is captured in group 1.
        string abbr = args.Match.Groups[1].Value.ToLowerInvariant();

        // Determine the full form.
        string fullForm = abbr switch
        {
            "e.g." => "for example",
            "i.e." => "that is",
            "etc." => "and so on",
            _ => abbr
        };

        // Preserve any trailing punctuation captured in the named group "punct".
        string trailingPunct = args.Match.Groups["punct"].Value;

        // Set the replacement text.
        args.Replacement = fullForm + trailingPunct;

        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Create a new document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Examples: e.g., i.e.; etc! Also see e.g. and i.e.");

        // Prepare the regular expression.
        // Group 1 captures the abbreviation, named group "punct" captures optional trailing punctuation.
        Regex regex = new Regex(@"\b(e\.g\.|i\.e\.|etc\.)\b(?<punct>[\.,;:!?\)]?)",
                                RegexOptions.IgnoreCase);

        // Set up replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replace operation.
        doc.Range.Replace(regex, string.Empty, options);

        // Save the resulting document.
        doc.Save("AbbreviationReplaced.docx");
    }
}
