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

        // Start building a table.
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

        // Apply the built‑in "Grid Table 5 Dark" style.
        // The style can be referenced by its name.
        table.StyleName = "Grid Table 5 Dark";

        // Optionally enable conditional formatting (first row as header, row banding, etc.).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "GridTable5Dark.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
