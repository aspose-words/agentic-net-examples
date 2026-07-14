using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        Table table = builder.StartTable();

        // ---------- Row 1 ----------
        // Cell 1 – first cell of a horizontal merge (colspan = 2).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header A (colspan=2)");

        // Cell 2 – part of the previous horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Cell 3 – normal cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Header B");

        builder.EndRow();

        // ---------- Row 2 ----------
        // Cell 1 – first cell of a vertical merge (rowspan = 2).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("RowSpan Cell");

        // Cell 2 – normal cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 2,2");

        // Cell 3 – first cell of a horizontal merge (colspan = 2).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged C (colspan=2)");

        // Cell 4 – part of the previous horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.EndRow();

        // ---------- Row 3 ----------
        // Cell 1 – continuation of the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Cell 2 – normal cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 3,2");

        // Cell 3 – normal cell (the second part of the previous horizontal merge is omitted because it is merged).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Cell 3,3");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as HTML. Aspose.Words will generate proper colspan and rowspan attributes.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        doc.Save("Table.html", saveOptions);
    }
}
