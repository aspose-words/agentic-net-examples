using System;
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

        // First cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Fixed width cell 1");

        // Second cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Fixed width cell 2");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Disable automatic column resizing while preserving the existing widths.
        table.AllowAutoFit = false;

        // Save the document to a file.
        doc.Save("TableAutoFitDisabled.docx");
    }
}
