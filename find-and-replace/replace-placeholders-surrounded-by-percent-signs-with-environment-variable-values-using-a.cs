using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class EnvVarReplacingCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The matched placeholder, e.g., %USERNAME%
        string placeholder = args.Match.Value;
        // Extract the environment variable name without the surrounding percent signs.
        string varName = placeholder.Trim('%');
        // Retrieve the environment variable value; use empty string if not defined.
        string envValue = Environment.GetEnvironmentVariable(varName) ?? string.Empty;
        // Replace the placeholder with the environment variable value.
        args.Replacement = envValue;
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Current user: %USERNAME%");
        builder.Writeln("Home directory: %USERPROFILE%");
        builder.Writeln("Path variable: %PATH%");
        builder.Writeln("Undefined variable: %NON_EXISTENT_VAR%");

        // Define a regex that matches placeholders surrounded by percent signs.
        Regex placeholderRegex = new Regex("%[^%]+%");

        // Set up the find-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new EnvVarReplacingCallback();

        // Perform the replacement. The replacement string parameter is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No placeholders were replaced. Expected at least one replacement.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optionally, write a simple confirmation to the console.
        Console.WriteLine($"Replaced {replacedCount} placeholder(s). Output saved to '{outputPath}'.");
    }
}
