using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

namespace WatermarkTableCellExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 4x4 table.
            Table table = new Table(doc);
            table.SetBorders(LineStyle.Single, 1.0, Color.Black);

            for (int i = 0; i < 4; i++)
            {
                Row row = new Row(doc);
                for (int j = 0; j < 4; j++)
                {
                    Cell cell = new Cell(doc);
                    cell.AppendChild(new Paragraph(doc));
                    cell.FirstParagraph.AppendChild(new Run(doc, $"R{i + 1}C{j + 1}"));
                    row.Cells.Add(cell);
                }
                table.Rows.Add(row);
            }

            // Insert the table into the document body.
            doc.FirstSection.Body.AppendChild(table);

            // Define the top‑left cell of the spanning area (row 2, column 2 in 1‑based indexing).
            Cell spanningCell = table.Rows[1].Cells[1];

            // Merge cells to span rows 2‑3 and columns 2‑3.
            spanningCell.CellFormat.HorizontalMerge = CellMerge.First;
            spanningCell.CellFormat.VerticalMerge = CellMerge.First;

            Cell topRight = table.Rows[1].Cells[2];
            topRight.CellFormat.HorizontalMerge = CellMerge.Previous;
            topRight.CellFormat.VerticalMerge = CellMerge.First;

            Cell bottomLeft = table.Rows[2].Cells[1];
            bottomLeft.CellFormat.HorizontalMerge = CellMerge.First;
            bottomLeft.CellFormat.VerticalMerge = CellMerge.Previous;

            Cell bottomRight = table.Rows[2].Cells[2];
            bottomRight.CellFormat.HorizontalMerge = CellMerge.Previous;
            bottomRight.CellFormat.VerticalMerge = CellMerge.Previous;

            // Create a rectangle shape to act as a watermark inside the merged cell.
            Shape watermarkShape = new Shape(doc, ShapeType.Rectangle);
            watermarkShape.Width = 200;
            watermarkShape.Height = 100;
            watermarkShape.WrapType = WrapType.None;
            watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
            watermarkShape.FillColor = Color.FromArgb(128, Color.LightGray); // Semi‑transparent.
            watermarkShape.StrokeColor = Color.Gray;

            // Add centered text to the shape.
            Paragraph shapeParagraph = new Paragraph(doc);
            Run run = new Run(doc, "CONFIDENTIAL");
            run.Font.Size = 24;
            run.Font.Color = Color.Red;
            shapeParagraph.AppendChild(run);
            shapeParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            watermarkShape.AppendChild(shapeParagraph);

            // Insert the shape into the merged cell.
            spanningCell.FirstParagraph.AppendChild(watermarkShape);

            // Save the document.
            string outputPath = "WatermarkTableCell.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to {outputPath}");
            }
        }
    }
}
