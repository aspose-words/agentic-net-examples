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

        // Ensure both documents have no existing revisions (required before comparison).
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
        {
            docOriginal.Revisions.RejectAll();
            docEdited.Revisions.RejectAll();
        }

        // Set up comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Example options – adjust as needed.
            CompareMoves               = true,                     // Track moved text.
            IgnoreFormatting           = true,                     // Ignore formatting changes.
            IgnoreComments             = true,                     // Ignore comment changes.
            IgnoreTables               = false,                    // Include table changes.
            IgnoreFields               = false,                    // Include field changes.
            IgnoreFootnotes            = false,                    // Include footnote/endnote changes.
            IgnoreTextboxes            = false,                    // Include textbox changes.
            IgnoreHeadersAndFooters    = false,                    // Include header/footer changes.
            Granularity                = Granularity.WordLevel,    // Track changes at word level.
            Target                     = ComparisonTargetType.New // Use the edited document as the base.
        };

        // Perform the comparison. Revisions are added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Save the resulting document (with revisions) as DOCX.
        docOriginal.Save("ComparisonResult.docx");
    }
}
