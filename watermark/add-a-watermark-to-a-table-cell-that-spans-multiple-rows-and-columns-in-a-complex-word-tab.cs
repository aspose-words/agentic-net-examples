using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 4x4 table.
        Table table = builder.StartTable();
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Merge the top‑left four cells to create a cell that spans two rows and two columns.
        Cell startCell = table.Rows[0].Cells[0];
        startCell.CellFormat.HorizontalMerge = CellMerge.First;
        startCell.CellFormat.VerticalMerge = CellMerge.First;

        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;

        Cell bottomRightCell = table.Rows[1].Cells[1];
        bottomRightCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        bottomRightCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Move the cursor to the merged cell.
        builder.MoveToCell(0, 0, 0, 0); // tableIndex, rowIndex, columnIndex, cellIndex

        // Insert a rectangle shape that will act as a watermark inside the cell.
        Shape watermark = builder.InsertShape(ShapeType.Rectangle, 200, 50);
        watermark.WrapType = WrapType.None;               // No text wrapping.
        watermark.BehindText = true;                      // Appear behind cell content.
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        watermark.FillColor = Color.LightGray;            // Light background.
        watermark.StrokeColor = Color.Gray;               // Border color.

        // Add visible text inside the shape.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(new Run(doc, "CONFIDENTIAL"));
        watermark.AppendChild(para);

        // Position the shape to fill the merged cell.
        watermark.Left = 0;
        watermark.Top = 0;

        // Save the document.
        string outputPath = "WatermarkedTableCell.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Document saved successfully to '{outputPath}'."
            : "Failed to save the document.");
    }
}
