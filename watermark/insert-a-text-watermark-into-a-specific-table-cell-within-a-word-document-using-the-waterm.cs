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

        // Insert a 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // Choose the target cell (first row, first column).
        Cell targetCell = table.FirstRow.FirstCell;

        // Create a shape that will act as a text watermark inside the cell.
        Shape watermarkShape = new Shape(doc, ShapeType.Rectangle);
        watermarkShape.Width = 200;
        watermarkShape.Height = 50;
        watermarkShape.WrapType = WrapType.None;
        // Transparent fill so only the text is visible.
        watermarkShape.FillColor = Color.Transparent;
        watermarkShape.StrokeColor = Color.Empty;

        // Add the watermark text to the shape.
        Paragraph textParagraph = new Paragraph(doc);
        Run textRun = new Run(doc, "CONFIDENTIAL");
        textParagraph.AppendChild(textRun);
        watermarkShape.AppendChild(textParagraph);

        // Insert the shape into the first paragraph of the target cell.
        Paragraph cellParagraph = targetCell.FirstParagraph;
        cellParagraph.AppendChild(watermarkShape);

        // Save the document.
        string outputPath = "CellWatermark.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + Path.GetFullPath(outputPath));
        }
    }
}
