using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to the source documents.
            string originalPath = @"C:\Docs\Original.docx";
            string editedPath   = @"C:\Docs\Edited.docx";

            // Load the two documents that will be compared.
            Document docOriginal = new Document(originalPath);
            Document docEdited   = new Document(editedPath);

            // Configure comparison options.
            CompareOptions compareOptions = new CompareOptions
            {
                // Example: track changes at the word level and compare moves.
                Granularity = Granularity.WordLevel,
                CompareMoves = true,

                // Set which document is used as the base for comparison.
                // Current = docOriginal, New = docEdited.
                Target = ComparisonTargetType.Current
            };

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

            // Save the result (docOriginal now contains revision marks).
            string resultPath = @"C:\Docs\ComparedResult.docx";
            docOriginal.Save(resultPath);
        }
    }
}
