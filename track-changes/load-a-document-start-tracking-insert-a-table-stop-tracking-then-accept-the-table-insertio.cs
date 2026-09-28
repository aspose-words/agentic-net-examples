using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document and save it locally.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Initial content.");
        string filePath = "sample.docx";
        doc.Save(filePath);

        // -----------------------------------------------------------------
        // 2. Load the saved document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(filePath);

        // -----------------------------------------------------------------
        // 3. Enable track changes.
        // -----------------------------------------------------------------
        loadedDoc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // -----------------------------------------------------------------
        // 4. Insert a table while tracking is active – this creates a revision.
        // -----------------------------------------------------------------
        DocumentBuilder tbuilder = new DocumentBuilder(loadedDoc);
        tbuilder.StartTable();
        tbuilder.InsertCell();
        tbuilder.Write("Cell 1");
        tbuilder.EndRow();
        tbuilder.EndTable();

        // -----------------------------------------------------------------
        // 5. Stop tracking.
        // -----------------------------------------------------------------
        loadedDoc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // 6. Accept only the table insertion revision.
        //    The revision's ParentNode will be the inserted Table node.
        // -----------------------------------------------------------------
        foreach (Revision rev in loadedDoc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion &&
                rev.ParentNode != null &&
                rev.ParentNode.NodeType == NodeType.Table)
            {
                rev.Accept();
            }
        }

        // -----------------------------------------------------------------
        // 7. Save the final document.
        // -----------------------------------------------------------------
        loadedDoc.Save("result.docx");
    }
}
