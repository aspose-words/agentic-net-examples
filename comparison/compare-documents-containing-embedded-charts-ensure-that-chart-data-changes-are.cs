using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareCharts
{
    static void Main()
    {
        // Load the two documents that contain embedded charts.
        Document docOriginal = new Document("Original.docx");
        Document docEdited   = new Document("Edited.docx");

        // The Compare method requires both documents to be revision‑free.
        if (docOriginal.Revisions.Count != 0 || docEdited.Revisions.Count != 0)
            throw new InvalidOperationException("Both documents must have no revisions before comparison.");

        // Configure comparison options.
        // All element types are compared (tables, textboxes, etc.) and we explicitly
        // keep the DrawingML unique ID comparison enabled so that changes in chart data
        // are detected as revisions.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New
        };
        compareOptions.AdvancedOptions.IgnoreDmlUniqueId = false; // ensure chart data changes are tracked

        // Perform the comparison. Revisions are added to the original document.
        docOriginal.Compare(docEdited, "ChartComparer", DateTime.Now, compareOptions);

        // Save the result – it now contains revisions for any chart data modifications.
        docOriginal.Save("ComparisonResult.docx");
    }
}
