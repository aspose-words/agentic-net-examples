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

        // Start a table with two rows and two columns.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        SetCellMargins(builder.CellFormat, 5);
        builder.Write("Cell 1,1");

        // First row, second cell.
        builder.InsertCell();
        SetCellMargins(builder.CellFormat, 5);
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        SetCellMargins(builder.CellFormat, 5);
        builder.Write("Cell 2,1");

        // Second row, second cell.
        builder.InsertCell();
        SetCellMargins(builder.CellFormat, 5);
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithCellMargins.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method to set all four padding (margin) values for a cell.
    private static void SetCellMargins(CellFormat format, double points)
    {
        format.TopPadding = points;
        format.BottomPadding = points;
        format.LeftPadding = points;
        format.RightPadding = points;
    }
}
