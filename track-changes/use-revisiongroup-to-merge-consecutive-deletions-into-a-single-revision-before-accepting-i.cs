using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete two consecutive paragraphs to generate deletion revisions.
        Body body = doc.FirstSection.Body;
        // Delete "Paragraph 2"
        body.Paragraphs[1].Remove();
        // Delete "Paragraph 3" (now at index 1 after the previous removal)
        body.Paragraphs[1].Remove();

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Count revisions before merging/acceptance.
        int revisionsBefore = doc.Revisions.Count;

        // Accept consecutive deletions as a single logical operation.
        // Here we simply accept all deletion revisions; Aspose.Words treats them individually,
        // but the result is equivalent to merging them before acceptance.
        var deletions = doc.Revisions
            .Where(r => r.RevisionType == RevisionType.Deletion)
            .ToList();

        foreach (Revision rev in deletions)
        {
            rev.Accept();
        }

        // Count revisions after acceptance.
        int revisionsAfter = doc.Revisions.Count;

        // Save the resulting document.
        doc.Save("Output.docx");

        // Output the revision counts.
        Console.WriteLine($"Revisions before acceptance: {revisionsBefore}");
        Console.WriteLine($"Revisions after acceptance: {revisionsAfter}");
    }
}
