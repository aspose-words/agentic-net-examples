using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeWordsComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the original and edited documents.
            string originalPath = @"C:\Docs\Original.docx";
            string editedPath   = @"C:\Docs\Edited.docx";

            // Load the documents.
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

                // Example: ignore changes in comments.
                IgnoreComments = true,

                // Example: ignore changes in tables.
                IgnoreTables = false,

                // Set the base document for comparison (show changes in the new document).
                Target = ComparisonTargetType.New
            };

            // Advanced options – for example, ignore DrawingML unique IDs.
            compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Save the resulting document with revisions.
            string resultPath = @"C:\Docs\ComparisonResult.docx";
            docOriginal.Save(resultPath);
        }
    }
}
