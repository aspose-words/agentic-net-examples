using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareAndSummarize
{
    static void Main()
    {
        // Paths to the documents to compare.
        string originalPath = @"C:\Docs\Original.docx";
        string revisedPath = @"C:\Docs\Revised.docx";

        // Load the documents.
        Document originalDoc = new Document(originalPath);
        Document revisedDoc = new Document(revisedPath);

        // Set up comparison options (customize as needed).
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: track changes at the word level and compare all element types.
            Granularity = Granularity.WordLevel,
            CompareMoves = true,
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. The revisions are added to the original document.
        originalDoc.Compare(revisedDoc, "Comparer", DateTime.Now, compareOptions);

        // Build a textual summary of the comparison results.
        StringBuilder summary = new StringBuilder();

        // Total number of revisions.
        summary.AppendLine($"Total revisions: {originalDoc.Revisions.Count}");

        // Number of revision groups (each group represents a contiguous change block).
        summary.AppendLine($"Revision groups: {originalDoc.Revisions.Groups.Count}");

        // Detailed list of each revision.
        foreach (Revision rev in originalDoc.Revisions)
        {
            // Get a short description of the changed text (if any).
            string changedText = rev.ParentNode?.GetText()?.Trim() ?? string.Empty;

            // Append revision type and affected text.
            summary.AppendLine($"{rev.RevisionType}: \"{changedText}\"");
        }

        // Create a new document to hold the summary.
        Document summaryDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(summaryDoc);

        // Write the summary into the document.
        builder.Writeln("Comparison Summary");
        builder.Writeln("-------------------");
        builder.Writeln(summary.ToString());

        // Save the summary document.
        string summaryPath = @"C:\Docs\ComparisonSummary.docx";
        summaryDoc.Save(summaryPath);
    }
}
