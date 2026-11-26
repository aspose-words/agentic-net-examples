using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header row and three data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Writeln("Header 1");
        builder.InsertCell(); builder.Writeln("Header 2");
        builder.InsertCell(); builder.Writeln("Header 3");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell(); builder.Writeln($"Row {i} Col 1");
            builder.InsertCell(); builder.Writeln($"Row {i} Col 2");
            builder.InsertCell(); builder.Writeln($"Row {i} Col 3");
            builder.EndRow();
        }

        builder.EndTable();

        // Set a custom background color for the header row.
        Row headerRow = table.FirstRow;
        foreach (Cell cell in headerRow.Cells)
        {
            // Apply shading to each cell in the header row.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
