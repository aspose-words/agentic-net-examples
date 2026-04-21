using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class RevisionGroupMergeExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs – each paragraph will become a separate Run.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete two consecutive paragraphs to generate deletion revisions.
        // First deletion.
        doc.FirstSection.Body.Paragraphs[1].Remove();
        // Second deletion (the original third paragraph is now at index 1).
        doc.FirstSection.Body.Paragraphs[1].Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that consecutive deletions are merged into a single RevisionGroup.
        if (doc.Revisions.Groups.Count != 1)
            throw new InvalidOperationException("Expected a single revision group after consecutive deletions.");

        RevisionGroup deletionGroup = doc.Revisions.Groups[0];
        if (deletionGroup.RevisionType != RevisionType.Deletion)
            throw new InvalidOperationException("The revision group is not of type Deletion.");

        // Collect revisions that belong to the merged deletion group.
        List<Revision> revisionsToAccept = doc.Revisions
            .Where(r => r.Group == deletionGroup)
            .ToList();

        // Accept each revision individually (iteration over a copy avoids collection modification errors).
        foreach (Revision rev in revisionsToAccept)
            rev.Accept();

        // After accepting, there should be no remaining revisions.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Save the resulting document.
        doc.Save("MergedDeletions.docx");
    }
}
