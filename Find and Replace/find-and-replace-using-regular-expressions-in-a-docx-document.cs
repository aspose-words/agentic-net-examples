using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsRegexReplaceDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the folder containing the document.
            string dataDir = @"C:\Data\";

            // Load an existing DOCX document.
            Document doc = new Document(dataDir + "Input.docx");

            // Example 1: Replace every sequence of digits with a paragraph break.
            // The pattern \d+ matches one or more digits.
            // The replacement string "&p" inserts a paragraph break.
            doc.Range.Replace(new Regex(@"\d+"), "&p");

            // Example 2: Replace words "gray" or "grey" with "lavender" using a regular expression.
            // The pattern "gr(a|e)y" captures both spellings.
            doc.Range.Replace(new Regex("gr(a|e)y"), "lavender");

            // Example 3: Use substitutions in the replacement string.
            // This example swaps the order of two captured groups.
            // Pattern captures two words separated by a space.
            Regex swapRegex = new Regex(@"(\w+)\s+(\w+)");
            FindReplaceOptions options = new FindReplaceOptions
            {
                // Enable substitution syntax like $1, $2 in the replacement string.
                UseSubstitutions = true,
                // Legacy mode must be disabled to support substitutions.
                LegacyMode = false
            };
            // Replacement swaps the two words: "$2 $1"
            doc.Range.Replace(swapRegex, "$2 $1", options);

            // Save the modified document.
            doc.Save(dataDir + "Output.docx");
        }
    }
}
