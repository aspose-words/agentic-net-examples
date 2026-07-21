using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // Create a sample document containing placeholders surrounded by percent signs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello %USERNAME% from %COMPUTERNAME%!");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that matches placeholders like %PLACEHOLDER%.
        Regex placeholderPattern = new Regex("%[A-Za-z0-9_]+%");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new EnvVarReplacer());

        // Perform the replacement. The callback supplies the actual replacement text.
        int replacedCount = loaded.Range.Replace(placeholderPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholders were replaced. Ensure the environment variables exist.");

        // Save the modified document.
        loaded.Save(outputPath);
    }

    // Custom callback that replaces each matched placeholder with the corresponding environment variable value.
    private class EnvVarReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match includes the surrounding percent signs, e.g., "%USERNAME%".
            string placeholder = args.Match.Value;
            // Extract the variable name without the percent signs.
            string variableName = placeholder.Trim('%');

            // Retrieve the environment variable value; use an empty string if not defined.
            string envValue = Environment.GetEnvironmentVariable(variableName) ?? string.Empty;

            // Set the replacement text.
            args.Replacement = envValue;

            return ReplaceAction.Replace;
        }
    }
}
