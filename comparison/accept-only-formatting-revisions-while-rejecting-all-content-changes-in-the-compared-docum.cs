using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with two separate runs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Write("Hello ");
        builderOriginal.Write("world.");
        builderOriginal.Writeln(); // End of paragraph.

        // Clone the original to create a revised version.
        Document revised = (Document)original.Clone(true);
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        Paragraph para = revised.FirstSection.Body.FirstParagraph;

        // 1. Apply a formatting change (make the first run bold) – creates a FormatChange revision.
        Run firstRun = para.Runs[0];
        firstRun.Font.Bold = true;

        // 2. Delete the second run ("world.") – creates a Deletion revision.
        Run secondRun = para.Runs[1];
        secondRun.Remove();

        // 3. Insert a new paragraph – creates an Insertion revision.
        builderRevised.Writeln("Inserted paragraph.");

        // Compare the original document with the revised one.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Accept only formatting revisions, reject all other content changes.
        Revision[] revisions = original.Revisions.ToArray(); // Copy to avoid collection modification issues.
        foreach (Revision rev in revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
                rev.Accept();   // Keep formatting changes.
            else
                rev.Reject();   // Discard insertions, deletions, etc.
        }

        // Verify that no revisions remain after processing.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("There should be no remaining revisions.");

        // Save the resulting document.
        original.Save("Result.docx");
    }
}
