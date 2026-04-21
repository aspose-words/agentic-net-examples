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

        // Start a table and add a single row with two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        // End the first row.
        builder.EndRow();

        // Add a second row (the row we will set to auto height).
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        // Set the HeightRule of the current row to Auto.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Verify that the HeightRule was set to Auto.
        if (table.Rows[1].RowFormat.HeightRule != HeightRule.Auto)
            throw new InvalidOperationException("Row height rule was not set to Auto.");

        // Save the document.
        string outputPath = "TableRowHeightAuto.docx";
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
