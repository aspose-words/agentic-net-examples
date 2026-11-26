using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Short");
        builder.InsertCell();
        builder.Writeln("Short");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Short");
        builder.InsertCell();
        builder.Writeln("Short");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Disable automatic resizing of the table.
        // In the current Aspose.Words version the AutoFitBehavior.Fixed enum value is not available,
        // so we use the AllowAutoFit property to achieve the same effect.
        table.AllowAutoFit = false;

        // Set fixed column widths (e.g., 100 points for each column).
        // Column widths are defined by the cells in the first row.
        Row firstRow = table.FirstRow;
        for (int i = 0; i < firstRow.Cells.Count; i++)
        {
            Cell cell = firstRow.Cells[i];
            // Use the PreferredWidth.FromPoints factory method to create a fixed width.
            cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        }

        // Add long text to a cell after fixing widths to demonstrate that the column does not expand.
        Cell targetCell = table.Rows[0].Cells[0];
        targetCell.FirstParagraph.AppendChild(new Run(doc,
            "This is a very long piece of text that would normally cause the column to expand if AutoFit were enabled."));

        // Save the document to disk.
        string outputPath = "TableAutoFitFixed.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
