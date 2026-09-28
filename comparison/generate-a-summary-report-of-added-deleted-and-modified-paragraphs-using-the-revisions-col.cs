using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class RevisionSummaryExample
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(original);
        builderOrig.Writeln("Paragraph 1");
        builderOrig.Writeln("Paragraph 2");
        builderOrig.Writeln("Paragraph 3");

        // Create the revised document:
        // - Paragraph 1 is removed (deletion).
        // - Paragraph 2 is changed (modification).
        // - Paragraph 4 is added (insertion).
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);
        builderRev.Writeln("Paragraph 2 modified");
        builderRev.Writeln("Paragraph 3");
        builderRev.Writeln("Paragraph 4 added");

        // Perform comparison. Revisions are stored in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after comparison.");
        }

        // Separate insertion and deletion revisions that affect paragraphs.
        List<Revision> insertions = new List<Revision>();
        List<Revision> deletions = new List<Revision>();

        foreach (Revision rev in original.Revisions)
        {
            // Only consider paragraph-level revisions.
            if (rev.ParentNode is Paragraph)
            {
                if (rev.RevisionType == RevisionType.Insertion)
                    insertions.Add(rev);
                else if (rev.RevisionType == RevisionType.Deletion)
                    deletions.Add(rev);
            }
        }

        // Determine modified paragraphs by pairing a deletion with a subsequent insertion.
        int modifiedCount = 0;
        // Use a copy of the insertion list to allow removal of matched items.
        List<Revision> remainingInsertions = new List<Revision>(insertions);

        foreach (Revision delRev in deletions)
        {
            Paragraph? deletedPara = delRev.ParentNode as Paragraph;
            if (deletedPara == null) continue;

            Revision? matchingInsert = null;
            foreach (Revision insRev in remainingInsertions)
            {
                Paragraph? insertedPara = insRev.ParentNode as Paragraph;
                if (insertedPara == null) continue;

                // If the inserted paragraph's previous sibling is the deleted paragraph,
                // treat this as a modification.
                if (insertedPara.PreviousSibling == deletedPara)
                {
                    matchingInsert = insRev;
                    break;
                }
            }

            if (matchingInsert != null)
            {
                modifiedCount++;
                remainingInsertions.Remove(matchingInsert);
            }
        }

        // After pairing, the leftovers are pure additions or deletions.
        int addedCount = remainingInsertions.Count;
        int deletedCount = deletions.Count - modifiedCount;

        // Build the summary report.
        string report = $"Revision Summary:{Environment.NewLine}" +
                        $"Added paragraphs   : {addedCount}{Environment.NewLine}" +
                        $"Deleted paragraphs : {deletedCount}{Environment.NewLine}" +
                        $"Modified paragraphs: {modifiedCount}{Environment.NewLine}";

        // Output to console.
        Console.WriteLine(report);

        // Save the report to a text file.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionSummary.txt");
        File.WriteAllText(reportPath, report);

        // Also save the compared document with revisions for inspection.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedDocument.docx");
        original.Save(comparedPath);
    }
}
