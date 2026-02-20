using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Layout;

class DocumentComparisonToPdf
{
    static void Main()
    {
        // Load the original and the edited documents.
        Document docOriginal = new Document("Original.docx");
        Document docEdited = new Document("Edited.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,
            // Show changes in the edited document (new document is the target).
            Target = ComparisonTargetType.New,
            // Do not ignore any element types.
            CompareMoves = true,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false
        };

        // Perform the comparison. Revisions will be added to docOriginal.
        docOriginal.Compare(docEdited, "Comparer", DateTime.Now, compareOptions);

        // Configure how revisions are rendered when saving to a fixed-page format (PDF).
        RevisionOptions revOptions = docOriginal.LayoutOptions.RevisionOptions;
        revOptions.InsertedTextColor = RevisionColor.Green;
        revOptions.InsertedTextEffect = RevisionTextEffect.Italic;
        revOptions.DeletedTextColor = RevisionColor.Red;
        revOptions.DeletedTextEffect = RevisionTextEffect.Bold;
        revOptions.ShowOriginalRevision = true;   // Show original text alongside revisions.
        revOptions.ShowRevisionMarks = true;     // Render revision marks.
        revOptions.ShowInBalloons = ShowInBalloons.Format; // Show format revisions in balloons.

        // Save the result as a PDF file. The PDF will contain the tracked changes.
        docOriginal.Save("ComparisonResult.pdf");
    }
}
