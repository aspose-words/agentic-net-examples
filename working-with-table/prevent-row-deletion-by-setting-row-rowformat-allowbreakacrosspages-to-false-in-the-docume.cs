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

        // Start a table.
        builder.StartTable();

        // Insert first row with two cells.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();

        // Insert second row with two cells.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the first row and prevent it from breaking across pages.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.AllowBreakAcrossPages = false;

        // Save the document.
        string outputPath = "PreventRowBreak.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the document and confirm the setting persisted.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Row loadedFirstRow = loadedTable.Rows[0];
        if (loadedFirstRow.RowFormat.AllowBreakAcrossPages)
            throw new InvalidOperationException("AllowBreakAcrossPages was not set to false as expected.");

        // Program completed successfully.
    }
}
