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

        // Start a table with three rows and two columns.
        Table table = builder.StartTable();

        // ---------- Row 1 ----------
        // First column – start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First; // Top cell of the merged range.
        builder.Write("Merged vertically across 3 rows");

        // Second column – regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        // ---------- Row 2 ----------
        // First column – continue the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed in merged cells after the first one.

        // Second column – regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        // ---------- Row 3 ----------
        // First column – continue the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second column – regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 3, Col 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
