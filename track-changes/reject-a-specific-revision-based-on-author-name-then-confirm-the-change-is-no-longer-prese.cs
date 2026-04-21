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
        builder.Write("Original text. ");

        // First author makes an insertion revision.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Write("John's addition. ");
        doc.StopTrackRevisions();

        // Second author makes another insertion revision.
        doc.StartTrackRevisions("Jane Smith", DateTime.Now);
        builder.Write("Jane's addition. ");
        doc.StopTrackRevisions();

        // Ensure revisions were created.
        if (doc.Revisions.Count != 2)
            throw new Exception("Expected 2 revisions before rejection.");

        // Define criteria to match revisions by author name.
        IRevisionCriteria janeCriteria = new AuthorCriteria("Jane Smith");

        // Reject all revisions authored by Jane Smith.
        int rejectedCount = doc.Revisions.Reject(janeCriteria);
        if (rejectedCount == 0)
            throw new Exception("No revisions were rejected; criteria may be incorrect.");

        // Validate that only John's revision remains.
        if (doc.Revisions.Count != 1)
            throw new Exception("Unexpected number of revisions after rejection.");

        // Validate that Jane's inserted text is no longer present in the document.
        string text = doc.GetText();
        if (text.Contains("Jane's addition."))
            throw new Exception("Jane's revision was not removed from the document.");

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Custom criteria that matches revisions based on the author name.
    private class AuthorCriteria : IRevisionCriteria
    {
        private readonly string _authorName;

        public AuthorCriteria(string authorName)
        {
            _authorName = authorName ?? throw new ArgumentNullException(nameof(authorName));
        }

        public bool IsMatch(Revision revision)
        {
            return revision.Author == _authorName;
        }
    }
}
