using System;
using Aspose.Words;

namespace AsposeWordsRevisionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert initial content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Original paragraph.");

            // Start tracking revisions with an author name.
            doc.StartTrackRevisions("John Doe");

            // Insert text that will be recorded as an insertion revision.
            builder.Writeln("This line is added while tracking revisions.");

            // Delete a run to generate a deletion revision.
            // For demonstration, remove the first run of the first paragraph.
            if (doc.FirstSection.Body.FirstParagraph.Runs.Count > 0)
                doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

            // Stop tracking further changes.
            doc.StopTrackRevisions();

            // Add more text that will NOT be tracked.
            builder.Writeln("This line is added after tracking stopped.");

            // At this point the document contains revisions.
            Console.WriteLine($"Document has revisions: {doc.HasRevisions}");
            Console.WriteLine($"Number of revisions: {doc.Revisions.Count}");

            // Iterate through revisions and display their details.
            for (int i = 0; i < doc.Revisions.Count; i++)
            {
                Revision rev = doc.Revisions[i];
                Console.WriteLine($"Revision {i + 1}: Type={rev.RevisionType}, Author={rev.Author}, Date={rev.DateTime}");
            }

            // Accept all revisions to produce the final document without change markers.
            doc.Revisions.AcceptAll();

            // Save the resulting document to a DOCX file.
            doc.Save("TrackedChangesResult.docx");
        }
    }
}
