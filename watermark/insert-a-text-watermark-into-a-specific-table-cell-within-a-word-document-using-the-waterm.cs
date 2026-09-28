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

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndTable();

        // Locate the first cell (row 0, column 0).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell targetCell = table.Rows[0].Cells[0];

        // Create a shape that will act as a text watermark inside the cell.
        Shape watermarkShape = new Shape(doc, ShapeType.Rectangle);
        watermarkShape.Width = 200;
        watermarkShape.Height = 50;
        watermarkShape.WrapType = WrapType.Inline; // Keep it within the text flow.
        watermarkShape.FillColor = Color.Transparent;
        watermarkShape.StrokeColor = Color.Transparent;

        // Add the watermark text to the shape.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "CONFIDENTIAL");
        run.Font.Size = 24;
        run.Font.Color = Color.Red;
        run.Font.Bold = true;
        para.AppendChild(run);
        watermarkShape.AppendChild(para);

        // Insert the shape into the target cell.
        targetCell.FirstParagraph.AppendChild(watermarkShape);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + Path.GetFullPath(outputPath));
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
