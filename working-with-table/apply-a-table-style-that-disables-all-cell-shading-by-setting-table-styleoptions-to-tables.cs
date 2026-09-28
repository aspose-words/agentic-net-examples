using System;
using System.Drawing;
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
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in style.
        table.Style = doc.Styles["Table Grid"];

        // Disable any style‑based shading options.
        table.StyleOptions = TableStyleOptions.None;

        // Ensure all cells have no shading.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.ForegroundPatternColor = Color.Empty;
                cell.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
            }
        }

        // Save the document.
        string outputPath = "TableNoShading.docx";
        doc.Save(outputPath);

        // Verify the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
