using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that will not be tracked.
        builder.Writeln("Initial content before tracking.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert a simple 2‑cell table while tracking is enabled.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept only the table insertion revision.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode.NodeType == NodeType.Table)
            {
                rev.Accept();
                break; // Table insertion is a single revision; exit after accepting.
            }
        }

        // Save the resulting document.
        doc.Save("TableRevisionAccepted.docx");
    }
}
