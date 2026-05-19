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

        // Start building a table.
        Table table = builder.StartTable();

        // First column – set explicit width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Header 1");

        // Second column – set explicit width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Header 2");

        // End the header row.
        builder.EndRow();

        // Add a data row with the same column widths.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Data 1");

        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Data 2");

        // End the data row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply fixed layout – disables AutoFit and removes any table preferred width.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FixedLayoutTable.docx");
        doc.Save(outputPath);
    }
}
