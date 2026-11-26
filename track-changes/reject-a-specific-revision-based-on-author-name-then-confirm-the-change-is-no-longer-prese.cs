using System;
using System.Collections.Generic;
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

            // Add initial content.
            builder.Writeln("Original paragraph.");

            // ---------- Create revisions by author "Alice" ----------
            doc.StartTrackRevisions("Alice", DateTime.Now);
            // Insert a line while tracking is enabled.
            builder.Writeln("Alice's inserted line.");
            doc.StopTrackRevisions();

            // ---------- Create revisions by author "Bob" ----------
            doc.StartTrackRevisions("Bob", DateTime.Now);
            // Delete the first paragraph while tracking is enabled.
            Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
            firstParagraph.Remove(); // This creates a deletion revision.
            doc.StopTrackRevisions();

            // Save the document with revisions for reference (optional).
            doc.Save("TrackChanges_WithRevisions.docx");

            // ---------- Reject revisions authored by "Alice" ----------
            // Collect revisions to reject first to avoid modifying the collection during iteration.
            List<Revision> revisionsToReject = new List<Revision>();
            foreach (Revision rev in doc.Revisions)
            {
                if (rev.Author == "Alice")
                    revisionsToReject.Add(rev);
            }

            foreach (Revision rev in revisionsToReject)
                rev.Reject();

            // Accept all remaining revisions (e.g., Bob's deletion) so they stay applied.
            doc.AcceptAllRevisions();

            // Verify that Alice's inserted line is no longer present.
            string docText = doc.GetText();
            if (docText.Contains("Alice's inserted line."))
                throw new Exception("Revision by Alice was not rejected correctly.");

            // Verify that Bob's deletion revision still applied (original paragraph removed).
            if (docText.Contains("Original paragraph."))
                throw new Exception("Bob's deletion revision was unexpectedly removed.");

            // Save the final document after rejection.
            doc.Save("TrackChanges_AfterRejection.docx");

            // Output confirmation.
            Console.WriteLine("Revisions by Alice rejected successfully. Final document saved.");
        }
    }
}
