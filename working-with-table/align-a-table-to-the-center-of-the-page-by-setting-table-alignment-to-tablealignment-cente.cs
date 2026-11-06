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

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");
        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 3");
        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 4");
        // End the second row.
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Align the table to the center of the page.
        table.Alignment = TableAlignment.Center;

        // Define output file path.
        string outputPath = "CenteredTable.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
