using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a simple text watermark for the whole document (uses Watermark class).
        doc.Watermark.SetText("Document Watermark");

        // Build a table with 2 rows and 3 columns.
        Table table = new Table(doc);
        // Set table borders for visibility.
        table.SetBorders(LineStyle.Single, 1.0, Color.Black);

        // Add rows and cells.
        for (int rowIdx = 0; rowIdx < 2; rowIdx++)
        {
            Row row = new Row(doc);
            for (int colIdx = 0; colIdx < 3; colIdx++)
            {
                Cell cell = new Cell(doc);
                // Add a paragraph with placeholder text.
                Paragraph para = new Paragraph(doc);
                Run run = new Run(doc, $"R{rowIdx + 1}C{colIdx + 1}");
                para.AppendChild(run);
                cell.AppendChild(para);
                row.AppendChild(cell);
            }
            table.AppendChild(row);
        }

        // Insert the table into the document body.
        doc.FirstSection.Body.AppendChild(table);

        // Insert a simple shape watermark into each cell of the first row.
        Row firstRow = table.FirstRow;
        foreach (Cell cell in firstRow.Cells)
        {
            // Create a rectangle shape that will act as a watermark inside the cell.
            Shape watermarkShape = new Shape(doc, ShapeType.Rectangle);
            watermarkShape.Width = 100;
            watermarkShape.Height = 30;
            watermarkShape.FillColor = Color.LightGray;
            watermarkShape.StrokeColor = Color.LightGray;
            watermarkShape.WrapType = WrapType.None;
            watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
            watermarkShape.HorizontalAlignment = HorizontalAlignment.Center;
            watermarkShape.VerticalAlignment = VerticalAlignment.Center;

            // Add the watermark text inside the shape.
            Paragraph shapeParagraph = new Paragraph(doc);
            Run shapeRun = new Run(doc, "Cell Watermark");
            shapeParagraph.AppendChild(shapeRun);
            watermarkShape.AppendChild(shapeParagraph);

            // Insert the shape into the first paragraph of the cell.
            Paragraph firstPara = cell.FirstParagraph;
            firstPara.AppendChild(watermarkShape);
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
        }
        else
        {
            Console.WriteLine("Failed to create the document.");
        }
    }
}
