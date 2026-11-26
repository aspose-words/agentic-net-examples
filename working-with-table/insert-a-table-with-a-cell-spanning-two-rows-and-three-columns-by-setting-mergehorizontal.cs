using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        builder.StartTable();

        // ---------- First Row ----------
        // Cell (0,0) – first cell of the merged region.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Writeln("Merged Cell");

        // Cell (0,1) – part of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Cell (0,2) – part of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cell (1,0) – continuation of both horizontal and vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Cell (1,1) – continuation of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Cell (1,2) – continuation of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Reset cell formatting to defaults for any further content.
        builder.CellFormat.ClearFormatting();

        // Save the document.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
