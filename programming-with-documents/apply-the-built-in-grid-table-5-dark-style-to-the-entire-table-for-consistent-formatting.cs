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

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;
        // Apply all style options (first row, first column, row bands, etc.).
        table.StyleOptions = TableStyleOptions.Default;

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "GridTable5Dark.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
