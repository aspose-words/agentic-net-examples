using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS IS A SAMPLE DOCUMENT.");
        builder.Writeln("IT CONTAINS UPPERCASE WORDS LIKE ASP.NET AND CSHARP.");
        builder.Writeln("MixedCase words stay unchanged.");

        // Save the original document (optional, just to illustrate the workflow).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Regex that matches whole words composed entirely of uppercase letters (2 or more characters).
        Regex upperCaseWordRegex = new Regex(@"\b[A-Z]{2,}\b");

        // Set up FindReplaceOptions with a custom callback that converts each match to title case.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new TitleCaseReplacer()
        };

        // Perform the replacement. The replacement string is ignored because the callback supplies the value.
        int replacedCount = doc.Range.Replace(upperCaseWordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Custom callback that converts each regex match to title case.
    private class TitleCaseReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Convert the matched word to lower case, then to title case.
            string lower = args.Match.Value.ToLowerInvariant();
            string title = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(lower);
            args.Replacement = title;
            return ReplaceAction.Replace;
        }
    }
}
