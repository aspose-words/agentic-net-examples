using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Set a different background color for the header row cells.
        Row headerRow = table.Rows[0];
        foreach (Cell cell in headerRow.Cells)
        {
            cell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
        }

        // Save the document.
        string outputPath = "TableWithHeaderStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Inform that the process completed.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
