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

        // Begin a table.
        builder.StartTable();

        // ---------- First Row ----------
        // First cell – this cell will be merged vertically across three rows.
        builder.InsertCell();
        builder.Writeln("Merged Cell");
        // Set the vertical merge flag on the top cell.
        Cell topCell = (Cell)builder.CurrentParagraph.ParentNode;
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        // Second cell of the first row.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();

        // ---------- Second Row ----------
        // First cell – part of the vertical merge.
        builder.InsertCell();
        builder.Writeln("Row 2, merged part");
        Cell secondRowFirstCell = (Cell)builder.CurrentParagraph.ParentNode;
        secondRowFirstCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell of the second row.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        // ---------- Third Row ----------
        // First cell – part of the vertical merge.
        builder.InsertCell();
        builder.Writeln("Row 3, merged part");
        Cell thirdRowFirstCell = (Cell)builder.CurrentParagraph.ParentNode;
        thirdRowFirstCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell of the third row.
        builder.InsertCell();
        builder.Writeln("Row 3, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine("Document created: " + Path.GetFullPath(outputPath));
    }
}
