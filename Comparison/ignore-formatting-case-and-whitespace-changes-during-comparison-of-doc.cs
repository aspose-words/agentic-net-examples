using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonExample
{
    public class DocumentComparer
    {
        /// <summary>
        /// Compares two DOCX files while ignoring formatting, case, and whitespace changes.
        /// The resulting document contains revisions that represent the differences.
        /// </summary>
        /// <param name="originalPath">Path to the original document.</param>
        /// <param name="modifiedPath">Path to the document to compare against.</param>
        /// <param name="resultPath">Path where the comparison result will be saved.</param>
        public static void CompareIgnoreFormattingCaseWhitespace(string originalPath, string modifiedPath, string resultPath)
        {
            // Load the two documents.
            Document original = new Document(originalPath);
            Document modified = new Document(modifiedPath);

            // Configure comparison options.
            CompareOptions compareOptions = new CompareOptions
            {
                // Ignore any formatting changes (bold, italics, font size, etc.).
                IgnoreFormatting = true,
                // Perform a case‑insensitive comparison.
                IgnoreCaseChanges = true
                // Whitespace changes are treated as formatting changes, so they are ignored as well.
            };

            // Perform the comparison. The revisions are added to the original document.
            original.Compare(modified, "Comparer", DateTime.Now, compareOptions);

            // Save the document that now contains revision marks.
            original.Save(resultPath);
        }
    }

    public class Program
    {
        // Entry point required for a console application.
        public static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: DocumentComparisonExample <original.docx> <modified.docx> <result.docx>");
                return;
            }

            string originalPath = args[0];
            string modifiedPath = args[1];
            string resultPath   = args[2];

            try
            {
                DocumentComparer.CompareIgnoreFormattingCaseWhitespace(originalPath, modifiedPath, resultPath);
                Console.WriteLine($"Comparison completed. Result saved to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during comparison: {ex.Message}");
            }
        }
    }
}
