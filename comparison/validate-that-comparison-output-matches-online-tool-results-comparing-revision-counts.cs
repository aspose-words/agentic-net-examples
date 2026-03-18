using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeWordsComparisonValidator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the documents to compare.
            string originalPath = @"C:\Docs\Original.docx";
            string editedPath   = @"C:\Docs\Edited.docx";

            // Expected results from the online comparison tool.
            int expectedRevisionCount = 5;
            RevisionType[] expectedTypes = new RevisionType[]
            {
                RevisionType.Insertion,
                RevisionType.Deletion,
                RevisionType.FormatChange,
                RevisionType.Moving,
                RevisionType.Insertion
            };

            // Perform validation.
            ValidateComparison(originalPath, editedPath, expectedRevisionCount, expectedTypes);
        }

        /// <summary>
        /// Loads two documents, compares them, and validates that the resulting revisions
        /// match the expected count and types.
        /// </summary>
        /// <param name="originalPath">Path to the original document.</param>
        /// <param name="editedPath">Path to the edited document.</param>
        /// <param name="expectedCount">Expected number of revisions.</param>
        /// <param name="expectedTypes">Expected revision types (order‑independent).</param>
        static void ValidateComparison(string originalPath, string editedPath, int expectedCount, RevisionType[] expectedTypes)
        {
            // Load the documents.
            Document docOriginal = new Document(originalPath);
            Document docEdited   = new Document(editedPath);

            // Ensure both documents are revision‑free before comparison.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Both documents must not contain revisions before comparison.");

            // Set up comparison options (default values are sufficient for a full comparison).
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: you can change any flag here if you need to ignore specific elements.
                // All flags are left as false to include every change.
                CompareMoves = false,
                IgnoreFormatting = false,
                IgnoreCaseChanges = false,
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Target = ComparisonTargetType.New
            };

            // Perform the comparison. Revisions are added to the original document.
            docOriginal.Compare(docEdited, "Validator", DateTime.Now, compareOptions);

            // Validate revision count.
            int actualCount = docOriginal.Revisions.Count;
            if (actualCount != expectedCount)
                throw new Exception($"Revision count mismatch. Expected: {expectedCount}, Actual: {actualCount}");

            // Validate revision types (order‑independent).
            var actualTypes = new List<RevisionType>();
            foreach (Revision rev in docOriginal.Revisions)
                actualTypes.Add(rev.RevisionType);

            var expectedSet = new HashSet<RevisionType>(expectedTypes);
            var actualSet   = new HashSet<RevisionType>(actualTypes);

            if (!expectedSet.SetEquals(actualSet))
                throw new Exception($"Revision types mismatch. Expected: [{string.Join(", ", expectedSet)}], Actual: [{string.Join(", ", actualSet)}]");

            // If we reach this point, the comparison matches the expected results.
            Console.WriteLine("Comparison validation succeeded.");
            Console.WriteLine($"Total revisions: {actualCount}");
            Console.WriteLine("Revision types:");
            foreach (var type in actualSet)
                Console.WriteLine($" - {type}");
        }
    }
}
