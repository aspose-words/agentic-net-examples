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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 4x4 table.
        Table table = builder.StartTable();

        // Row 1
        builder.InsertCell(); // (0,0) – will be merged
        builder.InsertCell(); // (0,1)
        builder.InsertCell(); // (0,2)
        builder.InsertCell(); // (0,3)
        builder.EndRow();

        // Row 2
        builder.InsertCell(); // (1,0) – part of merged cell
        builder.InsertCell(); // (1,1) – part of merged cell
        builder.InsertCell(); // (1,2)
        builder.InsertCell(); // (1,3)
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.InsertCell();
        builder.InsertCell();
        builder.InsertCell();
        builder.EndRow();

        // Row 4
        builder.InsertCell();
        builder.InsertCell();
        builder.InsertCell();
        builder.InsertCell();
        builder.EndRow();

        builder.EndTable();

        // Merge the top‑left cell across 2 rows and 2 columns.
        Cell startCell = table.Rows[0].Cells[0];
        startCell.CellFormat.HorizontalMerge = CellMerge.First;
        startCell.CellFormat.VerticalMerge = CellMerge.First;

        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;
        table.Rows[1].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
        table.Rows[1].Cells[1].CellFormat.VerticalMerge = CellMerge.Previous;

        // Move the builder into the merged cell.
        builder.MoveTo(startCell.FirstParagraph);

        // Create a shape that will act as a watermark inside the cell.
        Shape watermarkShape = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 200,
            Height = 100,
            WrapType = WrapType.None,
            BehindText = true,
            RelativeHorizontalPosition = RelativeHorizontalPosition.Column,
            RelativeVerticalPosition = RelativeVerticalPosition.Paragraph,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            FillColor = Color.LightGray,
            StrokeColor = Color.LightGray
            // TextBox property is read‑only; we do not assign to it.
        };

        // Add the watermark text to the shape.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "CONFIDENTIAL")
        {
            Font = { Size = 24, Color = Color.DarkGray }
        };
        para.AppendChild(run);
        watermarkShape.AppendChild(para);

        // Insert the shape into the cell.
        builder.InsertNode(watermarkShape);

        // Save the document.
        string outputPath = "WatermarkedTableCell.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + Path.GetFullPath(outputPath));
        }
    }
}
