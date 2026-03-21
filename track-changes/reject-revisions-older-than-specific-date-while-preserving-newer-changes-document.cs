using System;
using System.Threading;
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
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add initial content.
            builder.Writeln("Original paragraph.");

            // First revision (older).
            doc.StartTrackRevisions("Author", DateTime.Now);
            builder.Writeln("First revision – will be considered old.");
            doc.StopTrackRevisions();

            // Wait to create a time gap between revisions.
            Thread.Sleep(2000);

            // Second revision (newer).
            doc.StartTrackRevisions("Author", DateTime.Now);
            builder.Writeln("Second revision – will be preserved.");
            doc.StopTrackRevisions();

            // Define the cutoff date – revisions older than this will be rejected.
            // Set it to a point between the two revisions.
            DateTime cutoff = DateTime.Now.AddSeconds(-1);

            // Reject all revisions that match the custom criteria (i.e., older than the cutoff).
            doc.Revisions.Reject(new DateRevisionCriteria(cutoff));

            // Save the resulting document; newer revisions remain intact.
            doc.Save("OutputPreservingNewerRevisions.docx");
        }
    }
}
