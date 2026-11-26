using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

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

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set default cell margins (padding) to 2 points for every cell in the table.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.LeftPadding = 2;
                cell.CellFormat.RightPadding = 2;
                cell.CellFormat.TopPadding = 2;
                cell.CellFormat.BottomPadding = 2;
            }
        }

        // Save the document.
        string fileName = "TableDefaultCellMargin.docx";
        doc.Save(fileName);

        // Verify the file was created.
        if (!File.Exists(fileName))
            throw new Exception($"Failed to create output file: {fileName}");

        Console.WriteLine($"Document saved successfully: {Path.GetFullPath(fileName)}");
    }
}
