using System;
using System.IO;
using Aspose.Words;

public class TrackChangesTableExample
{
    public static void Main()
    {
        // Define file paths.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "result.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it (this will be the file we load).
        // -----------------------------------------------------------------
        Document initialDoc = new Document();
        DocumentBuilder initBuilder = new DocumentBuilder(initialDoc);
        initBuilder.Writeln("Document before tracking changes.");
        initialDoc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document we just created.
        // -----------------------------------------------------------------
        Document doc = new Document(originalPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 3. Start tracking revisions.
        // -----------------------------------------------------------------
        doc.StartTrackRevisions("Author", DateTime.Now);

        // -----------------------------------------------------------------
        // 4. Insert a table while tracking is active – this creates an insertion revision.
        // -----------------------------------------------------------------
        builder.Writeln(); // Ensure the table starts on a new line.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 4");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // -----------------------------------------------------------------
        // 5. Stop tracking revisions.
        // -----------------------------------------------------------------
        doc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // 6. Accept only the table insertion revision.
        // -----------------------------------------------------------------
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode.NodeType == NodeType.Table)
            {
                rev.Accept();
            }
        }

        // -----------------------------------------------------------------
        // 7. Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}
