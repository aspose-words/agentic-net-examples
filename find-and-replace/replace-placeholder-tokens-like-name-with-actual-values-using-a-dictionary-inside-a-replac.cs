using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ReplaceEvaluator : IReplacingCallback
{
    private readonly IDictionary<string, string> _values;

    public ReplaceEvaluator(IDictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // The regex captures the token name without the surrounding braces.
        var match = args.Match;
        var tokenName = match.Groups[1].Value;

        if (_values.TryGetValue(tokenName, out var replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            // If the token is not found, keep the original placeholder.
            args.Replacement = match.Value;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Define placeholder values.
        var placeholders = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Date", DateTime.Today.ToString("yyyy-MM-dd") }
        };

        // Create a sample document containing placeholders.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Your appointment is scheduled for {{Date}}.");
        templateDoc.Save("template.docx");

        // Load the document for processing.
        var doc = new Document("template.docx");

        // Prepare a regex that matches {{Token}} patterns.
        var tokenRegex = new Regex(@"\{\{(\w+)\}\}");

        // Set up find/replace options with the custom callback.
        var options = new FindReplaceOptions(new ReplaceEvaluator(placeholders));

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(tokenRegex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the result.
        doc.Save("output.docx");
    }
}
