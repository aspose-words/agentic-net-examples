using System;
using Aspose.Words;
using Aspose.Words.Comparing;

class ContractComparison
{
    static void Main()
    {
        // Paths to the original and revised contract documents.
        string originalPath = @"C:\Docs\OriginalContract.docx";
        string revisedPath  = @"C:\Docs\RevisedContract.docx";
        string resultPath   = @"C:\Docs\ContractComparisonResult.docx";

        // Load the two documents that need to be compared.
        Document docOriginal = new Document(originalPath);
        Document docRevised  = new Document(revisedPath);

        // Configure comparison options to ignore all formatting changes.
        CompareOptions compareOptions = new CompareOptions
        {
            // When true, differences in font, style, color, etc. are not reported as revisions.
            IgnoreFormatting = true,

            // Optional: keep other flags at their defaults (false) to capture content changes.
            CompareMoves = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,

            // Use the revised document as the base for comparison (mirrors Word's "Show changes in" option).
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. All formatting differences will be ignored.
        docOriginal.Compare(docRevised, "LegalTeam", DateTime.Now, compareOptions);

        // Save the document that now contains revision marks for the content differences.
        docOriginal.Save(resultPath);
    }
}
