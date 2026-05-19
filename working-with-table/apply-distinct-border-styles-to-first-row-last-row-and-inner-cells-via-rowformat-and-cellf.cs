using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class TableBorderExample
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder for constructing the table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // ---------- First Row ----------
        // Apply a double blue border to the entire first row.
        builder.RowFormat.ClearFormatting();
        builder.RowFormat.Borders.LineStyle = LineStyle.Double;
        builder.RowFormat.Borders.Color = Color.Blue;
        builder.RowFormat.Borders.LineWidth = 2.0;

        // Create three cells for the first row.
        builder.InsertCell();
        builder.Write("First Row, Cell 1");
        builder.InsertCell();
        builder.Write("First Row, Cell 2");
        builder.InsertCell();
        builder.Write("First Row, Cell 3");
        builder.EndRow();

        // ---------- Inner Rows ----------
        // Define a single green border for cells in inner rows.
        builder.CellFormat.ClearFormatting();
        builder.CellFormat.Borders.LineStyle = LineStyle.Single;
        builder.CellFormat.Borders.Color = Color.Green;
        builder.CellFormat.Borders.LineWidth = 1.0;

        // First inner row.
        builder.InsertCell();
        builder.Write("Inner Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Inner Row 1, Cell 2");
        builder.InsertCell();
        builder.Write("Inner Row 1, Cell 3");
        builder.EndRow();

        // Second inner row.
        builder.InsertCell();
        builder.Write("Inner Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Inner Row 2, Cell 2");
        builder.InsertCell();
        builder.Write("Inner Row 2, Cell 3");
        builder.EndRow();

        // ---------- Last Row ----------
        // Apply a dash‑dot red border to the entire last row.
        builder.RowFormat.ClearFormatting();
        builder.RowFormat.Borders.LineStyle = LineStyle.DashDotStroker; // Dash‑dot style
        builder.RowFormat.Borders.Color = Color.Red;
        builder.RowFormat.Borders.LineWidth = 1.5;

        // Create three cells for the last row.
        builder.InsertCell();
        builder.Write("Last Row, Cell 1");
        builder.InsertCell();
        builder.Write("Last Row, Cell 2");
        builder.InsertCell();
        builder.Write("Last Row, Cell 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBorders.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
