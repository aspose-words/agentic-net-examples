using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocuments
{
    static void Main()
    {
        // Load the original and edited documents from disk.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // Ensure both documents have no revisions before starting the comparison.
        if (docOriginal.Revisions.Count == 0 && docEdited.Revisions.Count == 0)
        {
            // Configure comparison options.
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
                Granularity                = Granularity.WordLevel
            };

            // Advanced options – ignore differences that are not relevant for the comparison.
            compareOptions.AdvancedOptions.IgnoreDmlUniqueId = true;
            compareOptions.AdvancedOptions.IgnoreStoreItemId = true;

            // Perform the comparison. The original document will receive revision marks.
            docOriginal.Compare(docEdited, "Author", DateTime.Now, compareOptions);
        }

        // Save the document that now contains the revision tracking information.
        docOriginal.Save("ComparisonResult.docx");
    }
}
