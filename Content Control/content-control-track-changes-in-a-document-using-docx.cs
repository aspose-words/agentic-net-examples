using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable tracking of revisions (track changes).
        doc.TrackRevisions = true;

        // Build the initial content of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the original paragraph.");          // Paragraph 1
        builder.Writeln("Paragraph to be edited.");                 // Paragraph 2

        // Insert a simple table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // ----- Perform edits after tracking is enabled -----

        // 1. Modify the text of the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count > 0)
            firstParagraph.Runs[0].Text = "This is the modified paragraph.";

        // 2. Delete the second paragraph.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];
        secondParagraph.Remove();

        // 3. Change the text inside the second cell of the table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell secondCell = table.Rows[0].Cells[1];
        if (secondCell.FirstParagraph.Runs.Count > 0)
            secondCell.FirstParagraph.Runs[0].Text = "Edited Cell 2";

        // Save the document; it will contain the tracked changes (revisions).
        doc.Save("TrackedChanges.docx");
    }
}
