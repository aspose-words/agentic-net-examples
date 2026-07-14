using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs that we will later delete.
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph 2.");
        builder.Writeln("Paragraph 3.");
        builder.Writeln("Paragraph 4.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete two consecutive paragraphs to generate two separate deletion revisions.
        // These deletions will be adjacent in the document, so they belong to the same RevisionGroup.
        Paragraph para2 = doc.FirstSection.Body.Paragraphs[1];
        Paragraph para3 = doc.FirstSection.Body.Paragraphs[2];
        para2.Remove(); // first deletion revision
        para3.Remove(); // second deletion revision

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point we have two deletion revisions.
        Console.WriteLine($"Revisions before merging: {doc.Revisions.Count}");

        // Find the revision group that contains the consecutive deletions.
        RevisionGroup deletionGroup = null;
        foreach (RevisionGroup group in doc.Revisions.Groups)
        {
            if (group.RevisionType == RevisionType.Deletion)
            {
                deletionGroup = group;
                break;
            }
        }

        if (deletionGroup == null)
            throw new InvalidOperationException("No deletion revision group found.");

        // Collect revisions that belong to the identified group.
        var revisionsToAccept = doc.Revisions
                                   .Where(r => r.Group == deletionGroup)
                                   .ToList();

        // Accept each revision in the group.
        foreach (Revision rev in revisionsToAccept)
        {
            rev.Accept();
        }

        // After accepting the grouped deletions, there should be no remaining revisions.
        Console.WriteLine($"Revisions after accepting the merged group: {doc.Revisions.Count}");

        // Save the resulting document.
        doc.Save("MergedDeletions.docx");
    }
}
