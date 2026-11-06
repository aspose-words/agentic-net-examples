using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with three cells in a single row.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.InsertCell();
        builder.Writeln("Cell 3");

        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row row = table.Rows[0];

        // Merge the three cells across the row using CellFormat.HorizontalMerge.
        Cell firstCell = row.Cells[0];
        Cell secondCell = row.Cells[1];
        Cell thirdCell = row.Cells[2];

        // Set merge flags: first cell is the start, others are previous.
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        thirdCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Set text for the merged cell.
        firstCell.Paragraphs.Clear();
        firstCell.Paragraphs.Add(new Paragraph(doc));
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Merged Cell"));

        // Save the document.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
