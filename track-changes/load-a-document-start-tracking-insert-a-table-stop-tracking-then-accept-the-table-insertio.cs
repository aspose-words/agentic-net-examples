using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary original document and the final result.
        const string originalPath = "original.docx";
        const string finalPath = "final.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it – this provides a file to load.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Initial content.");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Load the document we just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // -----------------------------------------------------------------
        // 3. Start tracking revisions.
        // -----------------------------------------------------------------
        loadedDoc.StartTrackRevisions("John Doe", DateTime.Now);

        // -----------------------------------------------------------------
        // 4. Insert a table while tracking is active – this creates an insertion revision.
        // -----------------------------------------------------------------
        loadedBuilder.StartTable();
        loadedBuilder.InsertCell();
        loadedBuilder.Write("Cell 1");
        loadedBuilder.InsertCell();
        loadedBuilder.Write("Cell 2");
        loadedBuilder.EndRow();
        loadedBuilder.EndTable();

        // -----------------------------------------------------------------
        // 5. Stop tracking revisions.
        // -----------------------------------------------------------------
        loadedDoc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // 6. Accept only the table insertion revision(s).
        // -----------------------------------------------------------------
        foreach (Revision rev in loadedDoc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode.NodeType == NodeType.Table)
            {
                rev.Accept();
            }
        }

        // -----------------------------------------------------------------
        // 7. Save the resulting document.
        // -----------------------------------------------------------------
        loadedDoc.Save(finalPath);
    }
}
