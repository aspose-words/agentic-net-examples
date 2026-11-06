using System;
using System.Drawing;
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
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set a custom background color for the header row using CellFormat.Shading.
        Row headerRow = table.FirstRow;
        foreach (Cell cell in headerRow.Cells)
        {
            // Apply background color to each header cell.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved and can be re-opened.
        if (System.IO.File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // No further action required.
        }
    }
}
