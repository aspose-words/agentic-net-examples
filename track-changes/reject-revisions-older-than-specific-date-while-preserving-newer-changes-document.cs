using System;
using Aspose.Words;

namespace RevisionFilteringExample
{
    // Custom criteria that matches revisions older than a specified cutoff date.
    public class DateRevisionCriteria : IRevisionCriteria
    {
        private readonly DateTime _cutoffDate;

        public DateRevisionCriteria(DateTime cutoffDate)
        {
            _cutoffDate = cutoffDate;
        }

        // Returns true if the revision's DateTime is earlier than the cutoff.
        public bool IsMatch(Revision revision)
        {
            return revision.DateTime < _cutoffDate;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the document that contains tracked changes.
            Document doc = new Document("InputWithRevisions.docx");

            // Define the cutoff date – revisions older than this will be rejected.
            DateTime cutoff = new DateTime(2023, 1, 1);

            // Reject all revisions that match the custom criteria (i.e., older than the cutoff).
            doc.Revisions.Reject(new DateRevisionCriteria(cutoff));

            // Save the resulting document; newer revisions remain intact.
            doc.Save("OutputPreservingNewerRevisions.docx");
        }
    }
}
