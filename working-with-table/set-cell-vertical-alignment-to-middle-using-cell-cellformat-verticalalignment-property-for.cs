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

        // First cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        // Set vertical alignment to middle (center) for the first cell.
        Cell cell1 = builder.CurrentParagraph.ParentNode as Cell;
        if (cell1 != null)
            cell1.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // Second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        // Set vertical alignment to middle (center) for the second cell.
        Cell cell2 = builder.CurrentParagraph.ParentNode as Cell;
        if (cell2 != null)
            cell2.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "VerticalAlignmentTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
