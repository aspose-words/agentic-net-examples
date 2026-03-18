using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class PrefixReplacer : IReplacingCallback
{
    // Prefix to prepend to each matched word.
    private readonly string _prefix;

    public PrefixReplacer(string prefix)
    {
        _prefix = prefix;
    }

    // This method is called for every match found by the replace operation.
    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Build the new replacement text by adding the prefix.
        args.Replacement = _prefix + args.Match.Value;
        // Instruct the engine to replace the current match.
        return ReplaceAction.Replace;
    }
}

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample text containing several words.
        builder.Writeln("Aspose.Words makes document processing easy and powerful.");

        // Set up find-and-replace options with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PrefixReplacer("PRE_");

        // Use a regular expression to match whole words.
        Regex wordPattern = new Regex(@"\b\w+\b");

        // Perform the replace operation. The replacement string is ignored because the callback provides it.
        doc.Range.Replace(wordPattern, string.Empty, options);

        // Save the modified document.
        doc.Save("PrefixedWords.docx");
    }
}
