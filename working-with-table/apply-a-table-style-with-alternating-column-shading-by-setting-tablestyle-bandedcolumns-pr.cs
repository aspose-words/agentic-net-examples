using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Add a few data rows.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable column banding (alternating column shading).
        table.StyleOptions = TableStyleOptions.ColumnBands;

        // Save the document.
        string outputPath = "TableWithColumnBanding.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
