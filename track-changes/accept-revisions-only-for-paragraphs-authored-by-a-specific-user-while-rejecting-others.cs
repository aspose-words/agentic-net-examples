using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Target author whose revisions we want to keep.
        const string targetAuthor = "John Doe";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("Original paragraph.");

        // ---------- Revisions by John Doe ----------
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Writeln("Paragraph added by John.");
        doc.StopTrackRevisions();

        // ---------- Revisions by Jane Smith ----------
        doc.StartTrackRevisions("Jane Smith", DateTime.Now);
        builder.Writeln("Paragraph added by Jane.");

        // Create a deletion revision (author Jane) by removing the first paragraph.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[0];
        paragraphToDelete.Remove();
        doc.StopTrackRevisions();

        // Accept only revisions authored by the target author.
        doc.Revisions.Accept(new RevisionCriteria(targetAuthor));

        // Reject all remaining revisions (those not authored by the target author).
        doc.Revisions.RejectAll();

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

// Criteria that matches revisions based on author name.
public class RevisionCriteria : IRevisionCriteria
{
    private readonly string _authorName;

    public RevisionCriteria(string authorName)
    {
        _authorName = authorName;
    }

    public bool IsMatch(Revision revision)
    {
        return revision.Author == _authorName;
    }
}
