using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin tracking changes – all subsequent edits will be recorded as revisions.
        doc.StartTrackRevisions("John Doe");

        // Insert a simple 1‑row, 2‑cell table while tracking is active.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Stop tracking – further edits will not be recorded as revisions.
        doc.StopTrackRevisions();

        // Accept only the insertion revision that represents the newly added table.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.ParentNode == table)
            {
                rev.Accept();
                break;
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
