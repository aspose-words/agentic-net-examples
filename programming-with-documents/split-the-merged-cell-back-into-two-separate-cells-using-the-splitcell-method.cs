using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with two columns.
        builder.StartTable();

        // Insert first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Insert second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the row.
        builder.EndRow();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Merge the two cells horizontally.
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[0].Cells[1];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---- Split the merged cell back into two separate cells ----
        // Aspose.Words does not provide a SplitCell method in older versions.
        // To "split" the merged cells, simply reset the merge flags.
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;

        // Write distinct text into each cell after splitting.
        firstCell.FirstParagraph.Runs.Clear();
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Split Cell A"));

        secondCell.FirstParagraph.Runs.Clear();
        secondCell.FirstParagraph.AppendChild(new Run(doc, "Split Cell B"));

        // Save the document.
        string outputPath = "SplitCell.docx";
        doc.Save(outputPath);
    }
}
