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

        // ---------- Header row with merged cells ----------
        // First cell – start of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Header");

        // Second cell – continues the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // Third cell – continues the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // End the header row.
        builder.EndRow();

        // ---------- Data row ----------
        // Reset merge setting for normal cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // First data cell.
        builder.InsertCell();
        builder.Writeln("Column 1");

        // Second data cell.
        builder.InsertCell();
        builder.Writeln("Column 2");

        // Third data cell.
        builder.InsertCell();
        builder.Writeln("Column 3");

        // End the data row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "MergedHeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
