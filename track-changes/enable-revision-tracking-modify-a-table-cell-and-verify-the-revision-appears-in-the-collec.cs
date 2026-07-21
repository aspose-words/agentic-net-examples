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

        // Build a simple 1x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Enable revision tracking.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Modify the text of the first cell – this change will be recorded as a revision.
        Cell firstCell = table.FirstRow.FirstCell;
        firstCell.RemoveAllChildren(); // clear existing content
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Updated Cell 1");
        para.AppendChild(run);
        firstCell.AppendChild(para);

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that a revision was created.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Ensure the revision corresponds to the modified cell text.
        bool revisionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.ParentNode != null && rev.ParentNode.GetText().Contains("Updated Cell 1"))
            {
                revisionFound = true;
                break;
            }
        }

        if (!revisionFound)
            throw new InvalidOperationException("Expected revision for the table cell change was not found.");

        // Save the document.
        doc.Save("RevisionsTable.docx");
    }
}
