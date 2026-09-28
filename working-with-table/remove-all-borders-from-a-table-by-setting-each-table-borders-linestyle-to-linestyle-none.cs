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

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Remove all borders by setting the borders of each row and each cell to None.
        foreach (Row row in table.Rows)
        {
            // Remove borders from the row.
            row.RowFormat.Borders.LineStyle = LineStyle.None;

            // Remove borders from each cell in the row.
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Borders.LineStyle = LineStyle.None;
            }
        }

        // Save the document.
        string outputPath = "NoBordersTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
