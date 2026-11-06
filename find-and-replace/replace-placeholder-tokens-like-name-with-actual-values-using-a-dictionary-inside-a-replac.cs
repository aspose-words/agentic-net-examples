using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class PlaceholderReplacer : IReplacingCallback
{
    private readonly Document _document;
    private readonly Dictionary<string, string> _values;

    public PlaceholderReplacer(Document document, Dictionary<string, string> values)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The matched placeholder, e.g., "{{Name}}"
        string placeholder = args.Match.Value;

        // Extract the token name without the surrounding braces.
        string token = placeholder.Trim('{', '}');

        if (_values.TryGetValue(token, out string replacement))
        {
            // Replace the placeholder with the replacement text.
            args.Replacement = replacement;
            return ReplaceAction.Replace;
        }

        // If no replacement is found, leave the original text unchanged.
        return ReplaceAction.Skip;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholder tokens.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Your appointment is scheduled for {{Date}}.");
        builder.Writeln("Thank you, {{Company}}.");

        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath);

        // Load the document for processing.
        Document doc = new Document(inputPath);

        // Define replacement values.
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Name", "Alice Johnson" },
            { "Date", "2024-01-15" },
            { "Company", "Acme Corp" }
        };

        // Set up the callback that will perform token replacement.
        var callback = new PlaceholderReplacer(doc, replacements);
        var options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Regex to find all placeholders of the form {{Token}}.
        Regex placeholderRegex = new Regex(@"\{\{[^}]+\}\}");

        // Perform the find-and-replace operation.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
