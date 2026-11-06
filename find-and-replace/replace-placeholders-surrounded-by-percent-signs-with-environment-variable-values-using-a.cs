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
        builder.Writeln("User: %USERNAME%");
        builder.Writeln("Home directory: %USERPROFILE%");
        builder.Writeln("Path variable: %PATH%");
        builder.Writeln("Undefined variable: %NON_EXISTENT%");

        // Define a regex that matches placeholders like %VAR_NAME%.
        Regex placeholderRegex = new Regex("%([^%]+)%", RegexOptions.Compiled);

        // Set up find-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new EnvironmentVariableReplacer()
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}

// Custom callback that replaces each placeholder with the corresponding environment variable value.
public class EnvironmentVariableReplacer : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // The match contains the full placeholder, e.g., %USERNAME%.
        Match match = args.Match;
        // Extract the variable name without the surrounding percent signs.
        string variableName = match.Groups[1].Value;

        // Retrieve the environment variable value; use empty string if not defined.
        string envValue = Environment.GetEnvironmentVariable(variableName) ?? string.Empty;

        // Set the replacement text.
        args.Replacement = envValue;

        // Indicate that the replacement should be performed.
        return ReplaceAction.Replace;
    }
}
