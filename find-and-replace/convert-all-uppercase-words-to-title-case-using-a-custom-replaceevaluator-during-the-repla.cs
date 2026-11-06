using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class UppercaseToTitleCaseCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Get the matched uppercase word.
        string match = args.Match.Value;
        if (string.IsNullOrEmpty(match))
            return ReplaceAction.Skip;

        // Convert to title case (first letter uppercase, rest lowercase).
        string titleCase = char.ToUpper(match[0]) + match.Substring(1).ToLower();

        // Set the replacement text.
        args.Replacement = titleCase;
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with uppercase words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("THIS IS A TEST DOCUMENT. SOME WORDS LIKE EXAMPLE, SAMPLE, AND CODE ARE UPPERCASE.");
        builder.Writeln("Mixed case words such as Hello and WORLD should be handled correctly.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches whole words consisting of two or more uppercase letters.
        Regex uppercaseWordRegex = new Regex(@"\b[A-Z]{2,}\b");

        // Set up the find-and-replace options with the custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new UppercaseToTitleCaseCallback()
        };

        // Perform the replacement. The replacement string argument is ignored because the callback provides the text.
        int replacedCount = loaded.Range.Replace(uppercaseWordRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
