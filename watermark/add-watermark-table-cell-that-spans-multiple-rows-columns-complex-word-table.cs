using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class TableCellWatermarkExample
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a complex table with a cell that spans multiple rows and columns.
        Table table = builder.StartTable();

        // First row
        // Cell (0,0) – start of merged area (spans 2 rows x 2 columns)
        builder.InsertCell();
        builder.CellFormat.Width = 200;
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // Cell (0,1) – part of the horizontal merge
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write(string.Empty);

        // Cell (0,2) – regular cell
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Regular Cell");
        builder.EndRow();

        // Second row
        // Cell (1,0) – part of the vertical merge
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(string.Empty);

        // Cell (1,1) – part of both merges
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(string.Empty);

        // Cell (1,2) – another regular cell
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Regular Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Locate the merged cell (row 0, column 0).
        Cell targetCell = table.Rows[0].Cells[0];
        targetCell.EnsureMinimum();

        // Move the builder cursor to the first paragraph of the target cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert a WordArt shape that will act as the watermark.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText)
        {
            TextPath = { Text = "CONFIDENTIAL", FontFamily = "Arial", Bold = true },
            RelativeHorizontalPosition = RelativeHorizontalPosition.Column,
            // RelativeVerticalPosition omitted – default works inside a cell
            Left = 0,
            Top = 0,
            Width = 200,
            Height = 50,
            Rotation = -45,
            FillColor = Color.LightGray,
            StrokeColor = Color.LightGray,
            WrapType = WrapType.None,
            BehindText = true
        };

        // Add the shape to the cell's first paragraph.
        targetCell.FirstParagraph.AppendChild(watermarkShape);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
