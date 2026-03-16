using System;
using Aspose.Words;

namespace RevisionFilterExample
{
    // Custom criteria that matches revisions based on the author name.
    // The constructor parameter 'accept' determines whether the criteria
    // should match revisions to be accepted (true) or to be rejected (false).
    public class AuthorCriteria : IRevisionCriteria
    {
        private readonly string _author;
        private readonly bool _accept;

        public AuthorCriteria(string author, bool accept)
        {
            _author = author;
            _accept = accept;
        }

        // Returns true if the revision matches the criteria.
        // For acceptance we match revisions whose Author equals the target author.
        // For rejection we match revisions whose Author does NOT equal the target author.
        public bool IsMatch(Revision revision)
        {
            return _accept ? revision.Author == _author : revision.Author != _author;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the document that contains tracked changes.
            Document doc = new Document("Input.docx");

            // Specify the author whose revisions should be kept.
            const string targetAuthor = "John Doe";

            // Accept all revisions authored by the target user.
            doc.Revisions.Accept(new AuthorCriteria(targetAuthor, true));

            // Reject all revisions authored by anyone else.
            doc.Revisions.Reject(new AuthorCriteria(targetAuthor, false));

            // Save the resulting document.
            doc.Save("Output.docx");
        }
    }
}
