using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Settings;

namespace DocumentComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source documents.
            // Replace with actual file system locations.
            string originalPath = @"MyDir\Original.docx";
            string editedPath   = @"MyDir\Edited.docx";

            // Load the original and edited documents.
            Document docOriginal = new Document(originalPath);
            Document docEdited   = new Document(editedPath);

            // Ensure both documents have no revisions before comparison.
            if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
                throw new InvalidOperationException("Documents must not contain revisions before comparison.");

            // Configure comparison options.
            CompareOptions compareOptions = new CompareOptions
            {
                // Track changes at the word level.
                Granularity = Granularity.WordLevel,

                // Do not ignore formatting changes.
                IgnoreFormatting = false,

                // Do not ignore case changes.
                IgnoreCaseChanges = false,

                // Include comments, tables, fields, footnotes, textboxes, headers/footers in the comparison.
                IgnoreComments = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,

                // Use the edited document as the base for comparison (equivalent to Word's "Show changes in New").
                Target = ComparisonTargetType.New,

                // Advanced options – for example, ignore DrawingML unique IDs.
                AdvancedOptions = { IgnoreDmlUniqueId = true }
            };

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Save the result to a DOCX file.
            string resultPath = @"ArtifactsDir\ComparedResult.docx";
            docOriginal.Save(resultPath);
        }
    }
}
