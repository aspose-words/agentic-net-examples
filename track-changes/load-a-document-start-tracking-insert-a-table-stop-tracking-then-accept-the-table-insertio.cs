using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsTrackChangesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some initial content (not tracked).
            builder.Writeln("Document before tracking changes.");

            // Start tracking revisions.
            doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

            // Insert a table while tracking is enabled.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.EndRow();
            builder.EndTable();

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Ensure that a revision was created.
            if (!doc.HasRevisions || doc.Revisions.Count == 0)
                throw new InvalidOperationException("No revisions were generated.");

            // Find the insertion revision that corresponds to the table.
            foreach (Revision rev in doc.Revisions)
            {
                if (rev.RevisionType == RevisionType.Insertion &&
                    rev.ParentNode != null &&
                    rev.ParentNode.NodeType == NodeType.Table)
                {
                    // Accept only this specific revision.
                    rev.Accept();
                    break;
                }
            }

            // Save the resulting document.
            doc.Save("TableRevisionAccepted.docx");
        }
    }
}
