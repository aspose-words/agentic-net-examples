using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsFindReplaceExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

            // Create a sample document with uppercase words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("THIS IS A TEST. Another line with UPPERCASE WORDS like EXAMPLE and SAMPLE.");
            builder.Writeln("MIXED case WORDS should stay unchanged.");
            doc.Save(inputPath);

            // Load the document we just created.
            Document loadedDoc = new Document(inputPath);

            // Define a regex that matches whole words consisting of only uppercase letters.
            Regex uppercaseWordRegex = new Regex(@"\b[A-Z]{2,}\b");

            // Set up find/replace options with a custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new UppercaseToTitleCaseReplacer()
            };

            // Perform the replacement. The replacement string is ignored because the callback supplies it.
            int replacedCount = loadedDoc.Range.Replace(uppercaseWordRegex, string.Empty, options);

            // Validate that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

            // Save the modified document.
            loadedDoc.Save(outputPath);

            // Output the result count (optional, not required for the task).
            Console.WriteLine($"Replaced {replacedCount} uppercase word(s).");
        }

        // Callback that converts each matched uppercase word to title case.
        private class UppercaseToTitleCaseReplacer : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                string original = args.Match.Value;
                if (!string.IsNullOrEmpty(original))
                {
                    // Convert to title case: first letter uppercase, the rest lowercase.
                    string titleCase = char.ToUpper(original[0], CultureInfo.InvariantCulture) +
                                       original.Substring(1).ToLower(CultureInfo.InvariantCulture);
                    args.Replacement = titleCase;
                }
                else
                {
                    args.Replacement = original;
                }

                return ReplaceAction.Replace;
            }
        }
    }
}
