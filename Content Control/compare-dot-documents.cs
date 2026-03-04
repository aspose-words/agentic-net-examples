using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class DocumentComparison
{
    static void Main()
    {
        // Load the two documents to be compared.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Documents must not contain revisions before a comparison is performed.
        if (docOriginal.Revisions.Count > 0 || docEdited.Revisions.Count > 0)
            throw new InvalidOperationException("Both documents must be revision‑free before comparison.");

        // Set up comparison options – adjust flags as needed.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves               = false,
            IgnoreFormatting           = false,
            IgnoreCaseChanges          = false,
            IgnoreComments             = false,
            IgnoreTables               = false,
            IgnoreFields               = false,
            IgnoreFootnotes            = false,
            IgnoreTextboxes            = false,
            IgnoreHeadersAndFooters    = false,
            Target                     = ComparisonTargetType.New,
            Granularity                = Granularity.WordLevel   // track changes by word
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "JD", DateTime.Now, compareOptions);

        // Save the resulting document that now contains revision marks.
        docOriginal.Save("ComparisonResult.docx");
    }
}
