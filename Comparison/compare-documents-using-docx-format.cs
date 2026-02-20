using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the two DOCX documents to be compared.
        Document docOriginal = new Document(@"C:\Docs\Original.docx");
        Document docEdited   = new Document(@"C:\Docs\Edited.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level (use CharLevel for character granularity).
            Granularity = Granularity.WordLevel,

            // Do not ignore formatting, case changes, comments, tables, fields, footnotes, textboxes, or headers/footers.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,

            // Do not compare moves (optional, set to true to include move tracking).
            CompareMoves = false,

            // Use the current document as the base for comparison.
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. Revisions (tracked changes) will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document with revisions to a new file.
        docOriginal.Save(@"C:\Docs\ComparisonResult.docx");
    }
}
