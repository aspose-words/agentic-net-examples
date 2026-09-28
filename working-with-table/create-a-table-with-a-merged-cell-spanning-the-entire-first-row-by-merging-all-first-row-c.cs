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

        // Define the number of columns for the table.
        int columnCount = 4;

        // Start the table.
        builder.StartTable();

        // ----- First row: merged cell spanning all columns -----
        // Insert the first cell and mark it as the start of a horizontal merge.
        builder.InsertCell();
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Header");

        // Insert the remaining cells of the first row and mark them as continuation of the merge.
        for (int i = 1; i < columnCount; i++)
        {
            builder.InsertCell();
            Cell mergedCell = (Cell)builder.CurrentParagraph.ParentNode;
            mergedCell.CellFormat.HorizontalMerge = CellMerge.Previous;
            // Content inside merged cells is optional; leaving it empty.
        }

        // End the first row.
        builder.EndRow();

        // ----- Second row: normal cells (optional, just to show table structure) -----
        for (int i = 0; i < columnCount; i++)
        {
            builder.InsertCell();
            builder.Write($"Row 2, Col {i + 1}");
        }
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
