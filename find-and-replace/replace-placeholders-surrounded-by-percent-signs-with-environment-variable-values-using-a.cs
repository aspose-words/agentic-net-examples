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
        // Set sample environment variables.
        Environment.SetEnvironmentVariable("NAME", "John Doe");
        Environment.SetEnvironmentVariable("CITY", "New York");

        // Create a simple document containing placeholders surrounded by percent signs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello %NAME%, welcome to %CITY%!");
        builder.Writeln("Your home directory is %HOME%.");

        // Configure find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new PlaceholderReplacer()
        };

        // Use a regular expression to locate %PLACEHOLDER% patterns.
        Regex placeholderPattern = new Regex("%[^%]+%");

        // Perform the replace operation. The callback supplies the actual replacement text.
        int replacementCount = doc.Range.Replace(placeholderPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Custom callback that replaces each %PLACEHOLDER% with the corresponding environment variable value.
    private class PlaceholderReplacer : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Extract the placeholder name without the surrounding percent signs.
            string placeholder = args.Match.Value;
            string variableName = placeholder.Trim('%');

            // Retrieve the environment variable value; use empty string if not defined.
            string envValue = Environment.GetEnvironmentVariable(variableName) ?? string.Empty;

            // Set the replacement text.
            args.Replacement = envValue;

            return ReplaceAction.Replace;
        }
    }
}
