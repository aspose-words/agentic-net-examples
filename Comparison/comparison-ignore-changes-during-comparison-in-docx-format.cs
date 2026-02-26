using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class Program
{
    static void Main()
    {
        // Load the original and the edited DOCX documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Configure comparison options to ignore all types of changes.
        CompareOptions compareOptions = new CompareOptions
        {
            // Ignore formatting differences.
            IgnoreFormatting = true,
            // Ignore case changes (case‑insensitive comparison).
            IgnoreCaseChanges = true,
            // Do not compare comments.
            IgnoreComments = true,
            // Do not compare tables.
            IgnoreTables = true,
            // Do not compare fields.
            IgnoreFields = true,
            // Do not compare footnotes and endnotes.
            IgnoreFootnotes = true,
            // Do not compare the contents of text boxes.
            IgnoreTextboxes = true,
            // Do not compare headers and footers.
            IgnoreHeadersAndFooters = true,
            // Do not generate move revisions.
            CompareMoves = false,
            // Use the edited document as the base for comparison.
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal,
        // but because all change types are ignored, the resulting document
        // will contain no visible revisions.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        docOriginal.Save("ComparisonResult.docx");
    }
}
