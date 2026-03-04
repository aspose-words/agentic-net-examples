using System;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace DocumentComparisonDemo
{
    public class Comparator
    {
        public static void CompareDocuments()
        {
            // Load the original and the edited documents (DOCX format).
            Document docOriginal = new Document("Original.docx");
            Document docEdited   = new Document("Edited.docx");

            // Configure comparison options.
            // Example: ignore formatting changes, track changes in comments, and use the edited document as the base.
            CompareOptions compareOptions = new CompareOptions
            {
                IgnoreFormatting = true,
                IgnoreComments = false,
                Target = ComparisonTargetType.New,
                CompareMoves = false,
                IgnoreCaseChanges = false,
                IgnoreTables = false,
                IgnoreFields = false,
                IgnoreFootnotes = false,
                IgnoreTextboxes = false,
                IgnoreHeadersAndFooters = false,
                Granularity = Granularity.WordLevel
            };

            // Perform the comparison. Revisions will be added to docOriginal.
            docOriginal.Compare(docEdited, "AuthorInitials", DateTime.Now, compareOptions);

            // Save the result as a DOCX file.
            docOriginal.Save("ComparedResult.docx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            Comparator.CompareDocuments();
        }
    }
}
