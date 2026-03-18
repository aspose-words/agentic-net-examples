using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class EnvironmentVariableReplacer : IReplacingCallback
{
    // This method is called for each match of the regex pattern.
    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // args.Match.Value includes the surrounding percent signs, e.g. "%PATH%".
        string placeholder = args.Match.Value;

        // Extract the variable name without the percent signs.
        // The regex guarantees the format %NAME%, so we can trim the first and last character.
        string variableName = placeholder.Substring(1, placeholder.Length - 2);

        // Retrieve the environment variable value.
        string envValue = Environment.GetEnvironmentVariable(variableName);

        // If the variable is not defined, keep the original placeholder.
        // Otherwise replace it with the environment variable value.
        args.Replacement = envValue ?? placeholder;

        // Continue with the replacement.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document and add some sample text containing placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("User: %USERNAME%");
        builder.Writeln("Home directory: %USERPROFILE%");
        builder.Writeln("Undefined variable: %NON_EXISTENT%");

        // Set up FindReplaceOptions with our custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new EnvironmentVariableReplacer();

        // Use a regular expression to find all substrings surrounded by percent signs.
        // The pattern captures any characters except a percent sign, between two percent signs.
        Regex placeholderRegex = new Regex("%([^%]+)%");

        // Perform the replace operation. The actual replacement text is supplied by the callback.
        doc.Range.Replace(placeholderRegex, string.Empty, options);

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
