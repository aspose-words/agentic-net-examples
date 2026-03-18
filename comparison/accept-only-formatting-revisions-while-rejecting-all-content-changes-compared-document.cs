using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Replacing; // not required but harmless

class Program
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Configure comparison options:
        // - Track formatting changes (IgnoreFormatting = false).
        // - Ignore all other element types to reduce noise.
        // - Do not compare moves.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = true,
            IgnoreTables = true,
            IgnoreFields = true,
            IgnoreFootnotes = true,
            IgnoreTextboxes = true,
            IgnoreHeadersAndFooters = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Reviewer", DateTime.Now, compareOptions);

        // Accept only formatting revisions, reject all other content revisions.
        foreach (Revision rev in docOriginal.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();   // keep formatting change
            else
                rev.Reject();   // discard insertion, deletion, move, etc.
        }

        // Save the resulting document with only formatting revisions applied.
        docOriginal.Save("Result.docx");
    }
}
