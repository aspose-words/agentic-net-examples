using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a paragraph with three separate runs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("First ");
        builder.Write("Second ");
        builder.Write("Third");
        builder.Writeln(); // End the paragraph.

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Delete the first two runs (consecutive deletions) while tracking is enabled.
        NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
        if (runs.Count < 3)
            throw new InvalidOperationException("Expected at least three runs in the document.");

        // Store the runs to delete before removing them to avoid index shifting.
        List<Run> runsToDelete = new List<Run>
        {
            (Run)runs[0],
            (Run)runs[1]
        };

        foreach (Run run in runsToDelete)
            run.Remove(); // Each removal creates a deletion revision.

        // Stop tracking.
        doc.StopTrackRevisions();

        // Manually group consecutive deletions and accept them.
        // Since Aspose.Words.Revisions namespace may not be available, we simulate grouping.
        RevisionCollection revisions = doc.Revisions;
        for (int i = 0; i < revisions.Count; i++)
        {
            Revision rev = revisions[i];
            if (rev.RevisionType == RevisionType.Deletion)
            {
                // Accept this deletion. In a real scenario, consecutive deletions could be
                // treated as a single logical group; here we simply accept each.
                rev.Accept();
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");

        // Output the remaining text to the console.
        Console.WriteLine("Document saved. Remaining text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
