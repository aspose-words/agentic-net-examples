using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders surrounded by percent signs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("User: %USERNAME%");
        builder.Writeln("Home directory: %USERPROFILE%");
        builder.Writeln("Path separator: %PATH_SEPARATOR%"); // This variable likely does not exist.

        // Save the original document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to perform find-and-replace.
        Document loaded = new Document(inputPath);

        // Set up the find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new PlaceholderReplacer()
        };

        // Regex to match placeholders like %VAR_NAME%.
        Regex placeholderPattern = new Regex("%[A-Za-z0-9_]+%");

        // Perform the replacement. The replacement string is ignored because the callback supplies the value.
        int replacedCount = loaded.Range.Replace(placeholderPattern, string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced. Ensure that the document contains %PLACEHOLDER% patterns.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that replaces each matched placeholder with the corresponding environment variable value.
    private class PlaceholderReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The matched placeholder, e.g., %USERNAME%
            string placeholder = args.Match.Value;

            // Extract the variable name without the surrounding percent signs.
            string variableName = placeholder.Trim('%');

            // Retrieve the environment variable value; if not found, replace with an empty string.
            string? envValue = Environment.GetEnvironmentVariable(variableName);
            args.Replacement = envValue ?? string.Empty;

            return ReplaceAction.Replace;
        }
    }
}
