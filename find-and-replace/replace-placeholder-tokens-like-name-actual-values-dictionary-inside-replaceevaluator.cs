using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document and add some text containing placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {{FirstName}} {{LastName}}, welcome to {{Company}}.");

        // Dictionary that maps placeholder names to their replacement values.
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "FirstName", "John" },
            { "LastName", "Doe" },
            { "Company", "Acme Corp" }
        };

        // Set up FindReplaceOptions with a custom callback that performs the replacement.
        FindReplaceOptions options = new FindReplaceOptions(new PlaceholderReplacer(values));

        // Regular expression that matches {{Placeholder}} patterns.
        Regex placeholderRegex = new Regex(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

        // Perform the replace operation. The actual replacement text is supplied by the callback.
        doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

// Implements IReplacingCallback to replace each matched placeholder with the value from the dictionary.
class PlaceholderReplacer : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public PlaceholderReplacer(IDictionary<string, string> values)
    {
        _values = values;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Group 1 of the match contains the placeholder name without the surrounding braces.
        string key = args.Match.Groups[1].Value;

        // If the dictionary contains a value for the key, use it; otherwise keep the original placeholder.
        if (_values.TryGetValue(key, out string replacement))
            args.Replacement = replacement;
        else
            args.Replacement = args.Match.Value;

        return ReplaceAction.Replace;
    }
}
