using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // First row with a merged cell spanning three columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Merged Header");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.EndRow();

        // Second row with regular cells.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string fileName = "MergedCellTable.docx";
        doc.Save(fileName);

        // Verify the file was created.
        if (!File.Exists(fileName))
            throw new Exception($"Failed to create the output file: {fileName}");

        // Output confirmation.
        Console.WriteLine($"Document saved successfully: {Path.GetFullPath(fileName)}");
    }
}
