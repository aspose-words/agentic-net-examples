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

        // Build a simple 3x4 table.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 3");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 3");
        builder.EndRow();

        // Fourth row.
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 2");
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 3");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in table style. Use the correct style name "Table Grid".
        table.Style = doc.Styles["Table Grid"];

        // Enable alternating column shading (banded columns).
        table.StyleOptions |= TableStyleOptions.ColumnBands;

        // Save the document to disk.
        string outputPath = "TableWithBandedColumns.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
