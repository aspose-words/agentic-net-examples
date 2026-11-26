using System;
using Aspose.Words;

namespace TrackChangesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add three separate paragraphs (sentences).
            builder.Writeln("This is the first sentence.");
            builder.Writeln("This is the second sentence.");
            builder.Writeln("This is the third sentence.");

            // Start tracking revisions.
            doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

            // Delete the second paragraph while tracking is enabled.
            Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];
            secondParagraph.Remove(); // Generates a deletion revision.

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Find a deletion revision.
            Revision deletionRevision = null;
            foreach (Revision rev in doc.Revisions)
            {
                if (rev.RevisionType == RevisionType.Deletion)
                {
                    deletionRevision = rev;
                    break;
                }
            }

            if (deletionRevision == null)
                throw new InvalidOperationException("Expected a deletion revision was not found.");

            // Accept the deletion revision individually.
            deletionRevision.Accept();

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
