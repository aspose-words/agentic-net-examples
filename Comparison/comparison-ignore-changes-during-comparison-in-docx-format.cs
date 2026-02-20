using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Saving;

namespace ComparisonExample
{
    class Program
    {
        static void Main()
        {
            // Load the original and the edited documents.
            Document original = new Document("Original.docx");
            Document edited = new Document("Edited.docx");

            // Configure comparison options to ignore all changes that should not affect the result.
            CompareOptions compareOptions = new CompareOptions
            {
                // Ignore formatting changes (bold, italic, font size, etc.).
                IgnoreFormatting = true,
                // Ignore changes in headers and footers.
                IgnoreHeadersAndFooters = true,
                // Ignore changes in footnotes and endnotes.
                IgnoreFootnotes = true,
                // Ignore changes in comments.
                IgnoreComments = true,
                // Ignore changes in tables.
                IgnoreTables = true,
                // Ignore changes inside text boxes.
                IgnoreTextboxes = true,
                // Ignore changes in fields (e.g., DATE, PAGE).
                IgnoreFields = true,
                // Do not compare moves (optional, can be set to false to keep move tracking).
                CompareMoves = false,
                // Use the original document as the base for comparison.
                Target = ComparisonTargetType.Current
            };

            // Perform the comparison. The result (revisions) will be stored in the original document.
            original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

            // Save the result as a DOCX file using default save options.
            original.Save("ComparisonResult.docx", SaveFormat.Docx);
        }
    }
}
