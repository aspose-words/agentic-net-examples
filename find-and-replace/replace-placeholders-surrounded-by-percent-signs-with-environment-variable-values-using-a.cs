using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for font handling

namespace AsposeWordsFindReplaceDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with placeholders surrounded by percent signs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello %USERNAME%!");
            builder.Writeln("Your home directory is %USERPROFILE%.");
            builder.Writeln("The OS version is %OS%.");
            // Save the input document (optional, just for demonstration).
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document to perform find-and-replace.
            Document loadedDoc = new Document(inputPath);

            // Define a regular expression that matches %PLACEHOLDER% patterns.
            Regex placeholderRegex = new Regex("%[A-Za-z0-9_]+%");

            // Set up the custom callback that replaces each placeholder with the corresponding environment variable.
            var replacer = new EnvironmentVariableReplacer();

            // Configure find/replace options to use the callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = replacer
            };

            // Perform the replace operation. The replacement string argument is ignored because the callback supplies the value.
            int replacedCount = loadedDoc.Range.Replace(placeholderRegex, string.Empty, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("No placeholders were replaced. Ensure that matching placeholders exist and environment variables are set.");

            // Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);

            // Output a simple confirmation.
            Console.WriteLine($"Replaced {replacedCount} placeholder(s). Modified document saved to '{outputPath}'.");
        }
    }

    // Custom callback that replaces %PLACEHOLDER% with the value of the corresponding environment variable.
    public class EnvironmentVariableReplacer : IReplacingCallback
    {
        // Optional: keep a log of processed placeholders.
        private readonly List<string> _processedPlaceholders = new List<string>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ensure the match is not null.
            if (args?.Match == null)
                return ReplaceAction.Skip;

            string placeholder = args.Match.Value; // e.g., %USERNAME%
            // Trim the surrounding percent signs to obtain the environment variable name.
            string variableName = placeholder.Trim('%');

            // Retrieve the environment variable value; if not found, replace with an empty string.
            string envValue = Environment.GetEnvironmentVariable(variableName) ?? string.Empty;

            // Set the replacement text.
            args.Replacement = envValue;

            // Record the operation (optional).
            _processedPlaceholders.Add($"{placeholder} => \"{envValue}\"");

            return ReplaceAction.Replace;
        }

        // Expose the log if needed.
        public IReadOnlyList<string> ProcessedPlaceholders => _processedPlaceholders;
    }
}
