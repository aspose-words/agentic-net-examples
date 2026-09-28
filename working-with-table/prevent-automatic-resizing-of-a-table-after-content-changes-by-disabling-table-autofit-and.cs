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

        // Build a simple 2‑column table.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Writeln("Short");

        // Second cell with long initial text.
        builder.InsertCell();
        builder.Writeln("This is a long piece of text that would normally cause the column to expand.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Disable automatic resizing (AutoFit) and fix column widths.
        // Set a fixed width for each column (e.g., 100 points).
        foreach (Cell cell in table.FirstRow.Cells)
        {
            cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        }
        // Apply the fixed‑column‑width behavior.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Add more text to the second cell after AutoFit has been disabled.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.FirstParagraph.AppendChild(new Run(doc, " Additional long text that would normally expand the column if AutoFit were enabled."));

        // Save the document.
        string outputPath = "TableAutoFitDisabled.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
