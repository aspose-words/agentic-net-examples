using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with uppercase words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS IS A TEST. SOME Words ARE UPPERCASE like NASA and FBI.");
        builder.Writeln("ANOTHER LINE WITH WORDS SUCH AS USA, UN, AND WHO.");

        // Define a callback that converts each matched uppercase word to title case.
        IReplacingCallback callback = new UppercaseToTitleCaseReplacer();

        // Set up find/replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions(callback);

        // Regex to match whole words consisting of two or more uppercase letters.
        Regex uppercaseWordPattern = new Regex(@"\b[A-Z]{2,}\b");

        // Perform the replacement. The replacement string is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(uppercaseWordPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

        // Save the modified document.
        doc.Save("output.docx");

        // Output the result count (optional, not required for the task).
        Console.WriteLine($"Replaced {replacedCount} uppercase word(s).");
    }

    // Callback that converts matched uppercase words to title case.
    private class UppercaseToTitleCaseReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string original = args.Match.Value;
            // Convert to title case: first letter uppercase, the rest lowercase.
            string titleCase = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(original.ToLowerInvariant());
            args.Replacement = titleCase;
            return ReplaceAction.Replace;
        }
    }
}
