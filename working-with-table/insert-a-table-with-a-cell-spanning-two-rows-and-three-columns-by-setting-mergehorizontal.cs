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

        // Start building the table.
        builder.StartTable();

        // ---------- First Row ----------
        // Cell that will span 2 rows and 3 columns.
        builder.InsertCell();
        builder.Writeln("Merged Cell");
        // Set the first cell of the merge.
        Cell mergedCell = builder.CurrentParagraph.ParentNode as Cell;
        mergedCell.CellFormat.HorizontalMerge = CellMerge.First;
        mergedCell.CellFormat.VerticalMerge = CellMerge.First;

        // Additional cells in the first row (columns 4 and 5).
        builder.InsertCell();
        builder.Writeln("Cell 4-1");
        builder.InsertCell();
        builder.Writeln("Cell 5-1");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cells that continue the merged area.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            // Set continuation of both horizontal and vertical merge.
            Cell continuationCell = builder.CurrentParagraph.ParentNode as Cell;
            continuationCell.CellFormat.HorizontalMerge = CellMerge.Previous;
            continuationCell.CellFormat.VerticalMerge = CellMerge.Previous;
        }

        // Additional cells in the second row (columns 4 and 5).
        builder.InsertCell();
        builder.Writeln("Cell 4-2");
        builder.InsertCell();
        builder.Writeln("Cell 5-2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, you could load the document again to ensure it is valid.
        Document loadedDoc = new Document(outputPath);
        // No further actions; the program ends here.
    }
}
