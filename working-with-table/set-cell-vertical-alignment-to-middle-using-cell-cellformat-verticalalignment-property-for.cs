using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 1, Row 1");

        // First row, second cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 1, Row 2");

        // Second row, second cell.
        builder.InsertCell();
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "VerticalAlignmentTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Reload the document and verify the vertical alignment of the first cell.
        Document loaded = new Document(outputPath);
        Cell firstCell = loaded.FirstSection.Body.Tables[0].Rows[0].Cells[0];
        if (firstCell.CellFormat.VerticalAlignment != CellVerticalAlignment.Center)
            throw new Exception("Cell vertical alignment was not set correctly.");
    }
}
