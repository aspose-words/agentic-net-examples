using System;
using System.IO;
using Aspose.Words;

namespace TrackChangesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add three sentences – these are normal edits, not tracked.
            builder.Writeln("Sentence one.");
            builder.Writeln("Sentence two to delete.");
            builder.Writeln("Sentence three.");

            // Start tracking revisions. All subsequent changes will be recorded.
            doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

            // Delete the second sentence while tracking is enabled.
            // The paragraph at index 1 corresponds to "Sentence two to delete."
            Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[1];
            paragraphToDelete.Remove();

            // Stop tracking revisions. Further edits will not be recorded.
            doc.StopTrackRevisions();

            // Find the deletion revision and accept it individually.
            foreach (Revision rev in doc.Revisions)
            {
                if (rev.RevisionType == RevisionType.Deletion)
                {
                    rev.Accept(); // Accept the deletion, permanently removing the paragraph.
                    break; // Only one deletion revision expected.
                }
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TrackChangesDemo.docx");
            doc.Save(outputPath);
        }
    }
}
