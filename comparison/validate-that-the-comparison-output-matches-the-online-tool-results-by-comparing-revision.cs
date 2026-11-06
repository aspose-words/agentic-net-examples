using System;
using System.IO;
using Aspose.Words;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Original text.");
        builderOriginal.Writeln("Second line.");

        // Create the revised document with clear differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Revised text.");          // Changed line (deletion + insertion)
        builderRevised.Writeln("Second line.");           // Unchanged line
        builderRevised.Writeln("New paragraph.");         // Insertion

        // Perform comparison.
        string author = "ComparisonExample";
        DateTime compareDate = DateTime.Now;
        original.Compare(revised, author, compareDate);

        // Inspect revisions.
        int totalRevisions = original.Revisions.Count;
        int insertionCount = 0;
        int deletionCount = 0;
        int formattingCount = 0;

        foreach (Revision rev in original.Revisions)
        {
            switch (rev.RevisionType)
            {
                case RevisionType.Insertion:
                    insertionCount++;
                    break;
                case RevisionType.Deletion:
                    deletionCount++;
                    break;
                case RevisionType.FormatChange:
                    formattingCount++;
                    break;
            }
        }

        // Expected counts based on the modifications made above.
        // Aspose.Words treats a changed line as one deletion and one insertion,
        // plus the new paragraph adds another insertion.
        const int expectedTotal = 4;          // 3 insertions (revised line + new paragraph + implicit insertion) + 1 deletion
        const int expectedInsertions = 3;
        const int expectedDeletions = 1;
        const int expectedFormatting = 0;

        // Validate revision counts.
        if (totalRevisions != expectedTotal ||
            insertionCount != expectedInsertions ||
            deletionCount != expectedDeletions ||
            formattingCount != expectedFormatting)
        {
            throw new InvalidOperationException(
                $"Revision validation failed. Expected total={expectedTotal}, insertions={expectedInsertions}, " +
                $"deletions={expectedDeletions}, formatting={expectedFormatting}. " +
                $"Actual total={totalRevisions}, insertions={insertionCount}, deletions={deletionCount}, formatting={formattingCount}.");
        }

        // Save the compared document.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedPath);

        // Create a simple text report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionReport.txt");
        using (StreamWriter writer = new StreamWriter(reportPath, false))
        {
            writer.WriteLine("Revision Summary Report");
            writer.WriteLine("-----------------------");
            writer.WriteLine($"Total revisions   : {totalRevisions}");
            writer.WriteLine($"Insertions        : {insertionCount}");
            writer.WriteLine($"Deletions         : {deletionCount}");
            writer.WriteLine($"Formatting changes: {formattingCount}");
        }

        // Output to console for quick verification.
        Console.WriteLine("Comparison completed successfully.");
        Console.WriteLine($"Compared document saved to: {comparedPath}");
        Console.WriteLine($"Revision report saved to: {reportPath}");
    }
}
