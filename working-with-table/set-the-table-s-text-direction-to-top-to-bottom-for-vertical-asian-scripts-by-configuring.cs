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

        // Start a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        // Set the cell text direction to top‑to‑bottom (downward).
        builder.CellFormat.Orientation = TextOrientation.Downward;
        builder.Write("セル 1 行 1");

        // First row, second cell.
        builder.InsertCell();
        builder.CellFormat.Orientation = TextOrientation.Downward;
        builder.Write("セル 2 行 1");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.CellFormat.Orientation = TextOrientation.Downward;
        builder.Write("セル 1 行 2");

        // Second row, second cell.
        builder.InsertCell();
        builder.CellFormat.Orientation = TextOrientation.Downward;
        builder.Write("セル 2 行 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableTextDirection.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
