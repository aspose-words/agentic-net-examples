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

        // Build a 2x2 table manually (no EndCell/EndRow methods).
        Table table = new Table(doc);
        for (int row = 0; row < 2; row++)
        {
            Row tableRow = new Row(doc);
            for (int col = 0; col < 2; col++)
            {
                Cell cell = new Cell(doc);
                // Add placeholder text to the cell.
                Paragraph para = new Paragraph(doc);
                Run run = new Run(doc, $"Cell {row + 1},{col + 1}");
                para.AppendChild(run);
                cell.AppendChild(para);
                tableRow.Cells.Add(cell);
            }
            table.Rows.Add(tableRow);
        }

        // Append the table to the document body (valid parent for a Table node).
        doc.FirstSection.Body.AppendChild(table);

        // Move the cursor to the first cell (row 1, column 1).
        Cell targetCell = table.Rows[0].Cells[0];
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert a rectangle shape into the cell.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Configure the shape to layout inside the cell.
        shape.IsLayoutInCell = true;

        // Make the shape visible.
        shape.FillColor = Color.LightBlue;
        shape.StrokeColor = Color.DarkBlue;

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
