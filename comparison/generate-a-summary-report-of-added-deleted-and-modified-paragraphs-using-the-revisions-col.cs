using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class RevisionSummaryExample
{
    public static void Main()
    {
        // Create the original document with several paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1: unchanged.");
        builder.Writeln("Paragraph 2: to be modified.");
        builder.Writeln("Paragraph 3: to be deleted.");
        builder.Writeln("Paragraph 4: unchanged.");

        // Save the original document (optional, just for reference).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Create the revised document containing added, deleted, and modified paragraphs.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Paragraph 1: unchanged.");
        builder2.Writeln("Paragraph 2: has been modified."); // Modified paragraph.
        builder2.Writeln("Paragraph 4: unchanged."); // Paragraph 3 deleted, Paragraph 4 moved up.
        builder2.Writeln("Paragraph 5: newly added."); // Added paragraph.

        // Perform the comparison. Revisions will be stored in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Save the compared document.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedPath);

        // Analyze revisions to produce a summary of added, deleted, and modified paragraphs.
        List<Revision> insertions = new List<Revision>();
        List<Revision> deletions = new List<Revision>();

        foreach (Revision rev in original.Revisions)
        {
            // Consider only paragraph revisions.
            if (rev.ParentNode is Paragraph)
            {
                if (rev.RevisionType == RevisionType.Insertion)
                {
                    insertions.Add(rev);
                }
                else if (rev.RevisionType == RevisionType.Deletion)
                {
                    deletions.Add(rev);
                }
            }
        }

        // Estimate modified paragraphs as paired insertion/deletion at the same location.
        int pairedCount = Math.Min(insertions.Count, deletions.Count);
        int modified = pairedCount;
        int added = insertions.Count - pairedCount;
        int deleted = deletions.Count - pairedCount;

        // Build the summary report.
        string report = $"Revision Summary:{Environment.NewLine}" +
                        $"Added paragraphs: {added}{Environment.NewLine}" +
                        $"Deleted paragraphs: {deleted}{Environment.NewLine}" +
                        $"Modified paragraphs: {modified}{Environment.NewLine}";

        // Save the report to a text file.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionSummary.txt");
        File.WriteAllText(reportPath, report);
    }
}
