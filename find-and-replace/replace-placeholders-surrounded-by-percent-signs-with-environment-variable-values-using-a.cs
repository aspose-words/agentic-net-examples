using System;
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
        builder.Writeln("Path: %PATH%");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Define a regex that matches %PLACEHOLDER% patterns.
        Regex placeholderRegex = new Regex("%[A-Za-z0-9_]+%");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new PlaceholderReplacer();

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(placeholderRegex, string.Empty, options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Optional: display the resulting text.
        Console.WriteLine("Replaced document text:");
        Console.WriteLine(loaded.GetText());
    }

    // Callback that replaces each %PLACEHOLDER% with the corresponding environment variable value.
    private class PlaceholderReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the placeholder name without the surrounding percent signs.
            string placeholder = args.Match.Value;
            string envVarName = placeholder.Trim('%');

            // Retrieve the environment variable value.
            string envValue = Environment.GetEnvironmentVariable(envVarName) ?? string.Empty;

            // Set the replacement text.
            args.Replacement = envValue;

            return ReplaceAction.Replace;
        }
    }
}
