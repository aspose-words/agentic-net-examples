using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package reference

namespace UppercaseToTitleCaseExample
{
    // Callback that converts each matched uppercase word to title case.
    public class UppercaseToTitleCaseReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Original matched text (e.g., "EXAMPLE")
            string original = args.Match.Value;

            // Convert to title case: first letter upper, the rest lower.
            string titleCase = char.ToUpper(original[0]) + original.Substring(1).ToLower();

            // Set the replacement text.
            args.Replacement = titleCase;

            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document with uppercase words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("THIS is a TEST document. ASP.NET CORE and C# are mentioned.");
            builder.Writeln("ANOTHER LINE WITH UPPERCASE WORDS LIKE EXAMPLE AND SAMPLE.");

            // Define a regex that matches whole words consisting of two or more uppercase letters.
            Regex uppercaseWordPattern = new Regex(@"\b[A-Z]{2,}\b");

            // Set up find-and-replace options with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new UppercaseToTitleCaseReplacer()
            };

            // Perform the replace operation. The replacement string is ignored because the callback sets it.
            int replacedCount = doc.Range.Replace(uppercaseWordPattern, string.Empty, options);

            // Ensure that at least one replacement occurred.
            if (replacedCount == 0)
                throw new InvalidOperationException("Expected at least one uppercase word to be replaced.");

            // Save the modified document.
            const string outputPath = "UppercaseToTitleCaseOutput.docx";
            doc.Save(outputPath);

            // Optional: write a simple confirmation to the console.
            Console.WriteLine($"Replacements performed: {replacedCount}");
            Console.WriteLine($"Modified document saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
