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
        // Cell (0,0) - top‑left cell that will span 2 rows and 3 columns.
        Cell cell00 = builder.InsertCell();
        cell00.CellFormat.HorizontalMerge = CellMerge.First;
        cell00.CellFormat.VerticalMerge = CellMerge.First;
        builder.Writeln("Merged 2x3");

        // Cells (0,1) and (0,2) - part of the horizontal merge.
        Cell cell01 = builder.InsertCell();
        cell01.CellFormat.HorizontalMerge = CellMerge.Previous;

        Cell cell02 = builder.InsertCell();
        cell02.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Cell (0,3) - regular cell.
        Cell cell03 = builder.InsertCell();
        builder.Writeln("Cell 4, Row 1");

        // End first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cells (1,0), (1,1), (1,2) - continuation of the merged area.
        Cell cell10 = builder.InsertCell();
        cell10.CellFormat.HorizontalMerge = CellMerge.Previous;
        cell10.CellFormat.VerticalMerge = CellMerge.Previous;

        Cell cell11 = builder.InsertCell();
        cell11.CellFormat.HorizontalMerge = CellMerge.Previous;
        cell11.CellFormat.VerticalMerge = CellMerge.Previous;

        Cell cell12 = builder.InsertCell();
        cell12.CellFormat.HorizontalMerge = CellMerge.Previous;
        cell12.CellFormat.VerticalMerge = CellMerge.Previous;

        // Cell (1,3) - regular cell.
        Cell cell13 = builder.InsertCell();
        builder.Writeln("Cell 4, Row 2");

        // End second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
