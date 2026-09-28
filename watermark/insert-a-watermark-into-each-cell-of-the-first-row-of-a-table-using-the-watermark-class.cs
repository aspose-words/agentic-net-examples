using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a table with 3 rows and 3 columns.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Populate the table.
        for (int rowIdx = 0; rowIdx < 3; rowIdx++)
        {
            Row row = new Row(doc);
            table.AppendChild(row);
            for (int colIdx = 0; colIdx < 3; colIdx++)
            {
                Cell cell = new Cell(doc);
                row.AppendChild(cell);
                Paragraph para = new Paragraph(doc);
                cell.AppendChild(para);
                Run run = new Run(doc, $"R{rowIdx + 1}C{colIdx + 1}");
                para.AppendChild(run);
            }
        }

        // Apply a document‑level text watermark using the native API.
        doc.Watermark.SetText("DOC WATERMARK");

        // Insert a simple shape watermark into each cell of the first row.
        Row firstRow = table.FirstRow;
        foreach (Cell cell in firstRow.Cells)
        {
            // Shape that will act as a watermark inside the cell.
            Shape watermarkShape = new Shape(doc, ShapeType.Rectangle);
            watermarkShape.Width = 200;
            watermarkShape.Height = 50;
            watermarkShape.Rotation = -40;
            watermarkShape.WrapType = WrapType.None;
            watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
            watermarkShape.VerticalAlignment = VerticalAlignment.Center;
            watermarkShape.HorizontalAlignment = HorizontalAlignment.Center;
            watermarkShape.BehindText = true;

            // Appearance.
            watermarkShape.Fill.Color = Color.LightGray;
            watermarkShape.Stroke.Color = Color.LightGray;

            // Text for the shape (avoid prohibited TextPath properties).
            watermarkShape.TextPath.Text = "CONFIDENTIAL";
            watermarkShape.TextPath.FontFamily = "Arial";
            watermarkShape.TextPath.Bold = true;
            watermarkShape.TextPath.Underline = false;

            // Insert the shape into the first paragraph of the cell.
            Paragraph cellParagraph = cell.FirstParagraph;
            cellParagraph.AppendChild(watermarkShape);
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
