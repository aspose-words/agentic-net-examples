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
        builder.StartTable();

        // ---- First row (merged cell) ----
        // Insert the first cell and mark it as the start of a horizontal merge.
        builder.InsertCell();
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Header spanning all columns");

        // Insert the second cell and mark it as a continuation of the merge.
        builder.InsertCell();
        Cell secondCell = (Cell)builder.CurrentParagraph.ParentNode;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged part.

        // End the first row.
        builder.EndRow();

        // ---- Second row (regular cells) ----
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not created.");
        }
    }
}
