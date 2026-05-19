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

        // Build a sample 3x4 table.
        Table table = builder.StartTable();

        // First row.
        for (int col = 0; col < 4; col++)
        {
            builder.InsertCell();
            builder.Write($"R1C{col + 1}");
        }
        builder.EndRow();

        // Second row.
        for (int col = 0; col < 4; col++)
        {
            builder.InsertCell();
            builder.Write($"R2C{col + 1}");
        }
        builder.EndRow();

        // Third row.
        for (int col = 0; col < 4; col++)
        {
            builder.InsertCell();
            builder.Write($"R3C{col + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table firstTable = doc.FirstSection.Body.Tables[0];

        // Get row and column counts.
        int rowCount = firstTable.Rows.Count;
        int columnCount = firstTable.Rows[0].Cells.Count;

        // Output the counts.
        Console.WriteLine($"Table has {rowCount} rows and {columnCount} columns.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableInfo.docx");
        doc.Save(outputPath);
    }
}
