using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world."); // baseline content.

        // Create the revised document with both formatting and content changes.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        // Same text but apply a formatting change (make it bold).
        builderRevised.Writeln("Hello world.");
        Paragraph para = revised.FirstSection.Body.FirstParagraph;
        if (para?.Runs.Count > 0)
        {
            para.Runs[0].Font.Bold = true; // formatting revision.
        }
        // Add a new paragraph – this is a content insertion revision.
        builderRevised.Writeln("Additional paragraph.");

        // Perform the comparison. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were created during comparison.");

        // Accept only formatting revisions, reject all other types.
        List<Revision> revisions = original.Revisions.Cast<Revision>().ToList();
        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();   // Keep the formatting change.
            else
                rev.Reject();   // Discard content insertions/deletions.
        }

        // After processing, there should be no remaining revisions.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Some revisions were not processed correctly.");

        // Save the final document.
        original.Save("Result.docx");
    }
}
