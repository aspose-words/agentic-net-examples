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

        // Build a simple 2x3 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the number of rows and columns.
        int rowCount = table.Rows.Count;
        int columnCount = table.Rows[0].Cells.Count;

        // Output the counts.
        Console.WriteLine($"Table has {rowCount} rows and {columnCount} columns.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableInfo.docx");
        doc.Save(outputPath);
    }
}
