using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("Original paragraph.");

        // First author makes a revision.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Writeln("John's inserted paragraph.");
        doc.StopTrackRevisions();

        // Second author makes a revision.
        doc.StartTrackRevisions("Jane Smith", DateTime.Now);
        builder.Writeln("Jane's inserted paragraph.");
        doc.StopTrackRevisions();

        // Ensure that revisions were created.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("No revisions were created.");

        // Reject revisions authored by "Jane Smith".
        int rejectedCount = doc.Revisions.Reject(new AuthorCriteria("Jane Smith"));

        // Verify that at least one revision was rejected (Jane's).
        if (rejectedCount < 1)
            throw new InvalidOperationException($"Expected to reject at least 1 revision, but rejected {rejectedCount}.");

        // Verify that no remaining revision is authored by Jane Smith.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Author == "Jane Smith")
                throw new InvalidOperationException("A revision by Jane Smith still exists after rejection.");
        }

        // Confirm the document text no longer contains Jane's paragraph.
        string text = doc.GetText();
        if (text.Contains("Jane's inserted paragraph."))
            throw new InvalidOperationException("Jane's revision was not removed from the document.");

        // Save the resulting document.
        doc.Save("Result.docx");

        // Output confirmation.
        Console.WriteLine("Revision by Jane Smith rejected successfully.");
        Console.WriteLine($"Remaining revisions: {doc.Revisions.Count}");
        Console.WriteLine("Document text:");
        Console.WriteLine(text);
    }

    // Custom criteria to match revisions by author name.
    private class AuthorCriteria : IRevisionCriteria
    {
        private readonly string _authorName;

        public AuthorCriteria(string authorName)
        {
            _authorName = authorName;
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author == _authorName;
        }
    }
}
