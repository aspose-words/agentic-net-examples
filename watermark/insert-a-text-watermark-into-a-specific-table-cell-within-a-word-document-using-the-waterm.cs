using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Save the initial document (optional, just for reference).
        string initialPath = "Sample.docx";
        doc.Save(initialPath);

        // Apply a text watermark to the document using the Watermark API.
        doc.Watermark.SetText("SAMPLE WATERMARK");

        // Retrieve the watermark shape that was added to the primary header.
        HeaderFooter header = doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary];
        Shape watermarkShape = (Shape)header.GetChildNodes(NodeType.Shape, true)[0];

        // Detach the shape from the header.
        watermarkShape.Remove();

        // Locate the target cell (second row, second column).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell targetCell = table.Rows[1].Cells[1];

        // Insert the watermark shape into the target cell.
        // Ensure the shape is inline so it fits inside the cell without layout issues.
        watermarkShape.WrapType = WrapType.Inline;
        targetCell.FirstParagraph.AppendChild(watermarkShape);

        // Save the resulting document.
        string outputPath = "WatermarkedCell.docx";
        doc.Save(outputPath);

        // Simple verification that the file was created.
        Console.WriteLine(File.Exists(outputPath)
            ? $"Watermark inserted and saved to '{outputPath}'."
            : "Failed to create the output document.");
    }
}
