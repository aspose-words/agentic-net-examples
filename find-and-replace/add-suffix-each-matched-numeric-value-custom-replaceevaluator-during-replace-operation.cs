using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some sample text that contains numeric values.
        builder.Writeln("The package weighs 12 kilograms and the box is 5 kilograms heavy.");

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions();
        // Use a custom callback that adds a suffix to each numeric match.
        options.ReplacingCallback = new NumericSuffixAdder(" kg");

        // Perform the replace operation using a regular expression that matches numbers.
        // The replacement string is empty because the actual replacement text is set in the callback.
        doc.Range.Replace(new Regex(@"\d+"), string.Empty, options);

        // Save the result (optional, demonstrates lifecycle usage).
        doc.Save("Output.docx");
    }
}

// Custom callback that appends a suffix to each numeric match.
class NumericSuffixAdder : IReplacingCallback
{
    private readonly string _suffix;

    public NumericSuffixAdder(string suffix)
    {
        _suffix = suffix;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // args.Match.Value contains the original numeric text.
        // Append the desired suffix and assign it to the Replacement property.
        args.Replacement = args.Match.Value + _suffix;

        // Continue with the replacement.
        return ReplaceAction.Replace;
    }
}
