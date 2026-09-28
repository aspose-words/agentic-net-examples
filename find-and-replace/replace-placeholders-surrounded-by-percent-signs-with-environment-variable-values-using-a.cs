using System;
using System.IO;
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
        builder.Writeln("Current user: %USERNAME%");
        builder.Writeln("Home directory: %USERPROFILE%");
        builder.Writeln("Path variable: %PATH%");
        builder.Writeln("Undefined variable: %NON_EXISTENT_VAR%");

        // Define a regex that matches placeholders like %PLACEHOLDER%.
        Regex placeholderRegex = new Regex("%[^%]+%");

        // Set up find-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EnvVarReplacingCallback()
        };

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Custom callback that replaces each placeholder with the corresponding environment variable value.
public class EnvVarReplacingCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The matched placeholder, e.g., %USERNAME%
        string placeholder = args.Match.Value;

        // Extract the environment variable name without the surrounding percent signs.
        string variableName = placeholder.Trim('%');

        // Retrieve the environment variable value; use empty string if not defined.
        string envValue = Environment.GetEnvironmentVariable(variableName) ?? string.Empty;

        // Set the replacement text.
        args.Replacement = envValue;

        // Indicate that the match should be replaced.
        return ReplaceAction.Replace;
    }
}
