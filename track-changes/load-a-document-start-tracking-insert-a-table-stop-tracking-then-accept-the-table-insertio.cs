using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a blank document that will serve as the starting point.
        // -----------------------------------------------------------------
        const string initialPath = "initial.docx";
        var initialDoc = new Document();
        initialDoc.Save(initialPath);

        // -----------------------------------------------------------------
        // 2. Load the document and start tracking revisions.
        // -----------------------------------------------------------------
        var doc = new Document(initialPath);
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // -----------------------------------------------------------------
        // 3. Insert a table while tracking is enabled – this will generate
        //    insertion revisions.
        // -----------------------------------------------------------------
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before the table.");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // -----------------------------------------------------------------
        // 4. Stop tracking revisions.
        // -----------------------------------------------------------------
        doc.StopTrackRevisions();

        // -----------------------------------------------------------------
        // 5. Accept only the insertion revisions that were created for the
        //    table.  We copy the collection to an array first to avoid
        //    modifying the collection while iterating.
        // -----------------------------------------------------------------
        var revisions = doc.Revisions.Cast<Revision>().ToArray();

        foreach (var rev in revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion)
                rev.Accept();
        }

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        const string resultPath = "result.docx";
        doc.Save(resultPath);

        Console.WriteLine($"Document processed and saved to {resultPath}");
    }
}
