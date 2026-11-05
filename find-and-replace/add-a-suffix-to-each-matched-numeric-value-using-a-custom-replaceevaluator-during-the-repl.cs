using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class NumericSuffixReplacer : IReplacingCallback
{
    private readonly string _suffix;

    public NumericSuffixReplacer(string suffix) => _suffix = suffix;

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Append the suffix to the matched numeric value.
        string original = args.Match.Value;
        args.Replacement = original + _suffix;
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document with sample text containing numbers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Invoice numbers: 1001, 1002, and 1003.");
        builder.Writeln("Reference dates: 2021-09-30 and 2022-01-15.");

        // Define the suffix to add to each numeric match.
        const string suffix = "_SUF";

        // Set up find‑replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new NumericSuffixReplacer(suffix);

        // Regex pattern to match one or more digits.
        Regex regex = new Regex(@"\d+");

        // Perform the replace operation using the overload that accepts a regex,
        // a dummy replacement string (ignored by the callback), and the options.
        int replacedCount = doc.Range.Replace(regex, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one numeric replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
