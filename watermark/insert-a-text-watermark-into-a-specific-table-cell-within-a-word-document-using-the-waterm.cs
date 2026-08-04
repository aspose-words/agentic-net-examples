using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Output file path
        string outputPath = "WatermarkedTableCell.docx";

        // Create a new blank document
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple 2x2 table
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Move the cursor to the target cell (first table, first row, second column)
        // Parameters: tableIndex, rowIndex, columnIndex, cellIndex
        builder.MoveToCell(0, 0, 1, 0);

        // Insert a text watermark using the Watermark class
        // The watermark will appear on every page; placing the cursor in the cell
        // demonstrates that the operation can be performed after navigating to a cell.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document
        doc.Save(outputPath);

        // Simple validation that the file was created
        Console.WriteLine(File.Exists(outputPath) ? "Document saved successfully." : "Failed to save document.");
    }
}
