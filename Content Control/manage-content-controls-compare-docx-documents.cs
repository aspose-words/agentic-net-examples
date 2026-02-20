using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareDocumentsExample
{
    static void Main()
    {
        // Load the original and the revised documents.
        Document originalDoc = new Document(@"C:\Docs\Original.docx");
        Document revisedDoc = new Document(@"C:\Docs\Revised.docx");

        // Configure comparison options.
        CompareOptions compareOptions = new CompareOptions
        {
            // Track changes at the word level.
            Granularity = Granularity.WordLevel,

            // Do not ignore any element types.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreFootnotes = false,
            IgnoreHeadersAndFooters = false,
            IgnoreTables = false,
            IgnoreTextboxes = false,
            IgnoreFields = false,

            // Use the revised document as the base for comparison (equivalent to "Show changes in: New").
            Target = ComparisonTargetType.New,

            // Do not compare moved text as separate changes.
            CompareMoves = false
        };

        // Perform the comparison. Revisions will be added to the original document.
        originalDoc.Compare(revisedDoc, "Comparer", DateTime.Now, compareOptions);

        // Save the document with revisions.
        originalDoc.Save(@"C:\Docs\ComparisonResult.docx");
    }
}
