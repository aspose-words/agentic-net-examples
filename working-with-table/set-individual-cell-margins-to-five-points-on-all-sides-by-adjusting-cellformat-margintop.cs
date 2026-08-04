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

        // Start a table.
        Table table = builder.StartTable();

        // ---- First Row ----
        // First cell.
        Cell cell1 = builder.InsertCell();
        // Set padding (margin) of 5 points on all sides.
        cell1.CellFormat.SetPaddings(5, 5, 5, 5);
        builder.Write("Cell 1");

        // Second cell.
        Cell cell2 = builder.InsertCell();
        cell2.CellFormat.SetPaddings(5, 5, 5, 5);
        builder.Write("Cell 2");

        // End first row.
        builder.EndRow();

        // ---- Second Row ----
        // First cell.
        Cell cell3 = builder.InsertCell();
        cell3.CellFormat.SetPaddings(5, 5, 5, 5);
        builder.Write("Cell 3");

        // Second cell.
        Cell cell4 = builder.InsertCell();
        cell4.CellFormat.SetPaddings(5, 5, 5, 5);
        builder.Write("Cell 4");

        // End second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellMargins.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
