using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add three paragraphs (no tracking yet).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        builder.Writeln("Third paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Apply a style change to all existing paragraphs while tracking is enabled.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.Style = doc.Styles["Heading 1"];
        }

        // Make an additional change (insert a new paragraph) to guarantee a revision.
        builder.Writeln("Inserted while tracking.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Save the document (demonstrates that revisions are persisted).
        doc.Save("TrackedDocument.docx");

        // Verify that revisions were created.
        RevisionCollection revisions = doc.Revisions;
        if (revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Verify that there is exactly one revision group.
        if (revisions.Groups.Count != 1)
            throw new InvalidOperationException($"Expected 1 revision group, but found {revisions.Groups.Count}.");

        // Output verification result.
        Console.WriteLine("Revision group verification succeeded.");
    }
}
