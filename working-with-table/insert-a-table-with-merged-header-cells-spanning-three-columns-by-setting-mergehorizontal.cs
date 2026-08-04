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

        // Start the table.
        Table table = builder.StartTable();

        // ----- Header row (merged across three columns) -----
        // First cell: start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning three columns");

        // Second cell: merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell: merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the header row.
        builder.EndRow();

        // Reset merge settings for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ----- First data row (three separate cells) -----
        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.InsertCell();
        builder.Write("Row 1, Col 3");
        builder.EndRow();

        // ----- Second data row (three separate cells) -----
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.InsertCell();
        builder.Write("Row 2, Col 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "MergedHeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {Path.GetFullPath(outputPath)}");
        }
        else
        {
            throw new Exception("Failed to save the document.");
        }
    }
}
