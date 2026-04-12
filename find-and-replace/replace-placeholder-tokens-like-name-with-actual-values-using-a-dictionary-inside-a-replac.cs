using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ReplaceEvaluator : IReplacingCallback
{
    private readonly Dictionary<string, string> _values;

    public ReplaceEvaluator(Dictionary<string, string> values)
    {
        _values = values ?? throw new ArgumentNullException(nameof(values));
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // The matched token includes the braces, e.g. "{{Name}}".
        string token = args.Match.Value;
        // Extract the key between the braces.
        string key = token.Trim('{', '}');

        if (_values.TryGetValue(key, out string replacement))
        {
            args.Replacement = replacement;
        }
        else
        {
            // If the key is not found, keep the original token.
            args.Replacement = token;
        }

        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Your order from {{Company}} has been shipped.");
        builder.Writeln("Thank you for shopping with us!");

        // Define the placeholder values.
        var replacements = new Dictionary<string, string>
        {
            { "Name", "John Doe" },
            { "Company", "Acme Corp" }
        };

        // Configure find-and-replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ReplaceEvaluator(replacements)
        };

        // Use a regular expression to locate tokens like {{Key}}.
        int replacedCount = doc.Range.Replace(new Regex(@"\{\{(\w+)\}\}"), string.Empty, options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No placeholders were replaced.");
        }

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"Total placeholders replaced: {replacedCount}");
    }
}
