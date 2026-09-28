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

        // Build a simple table with 5 rows and 2 columns.
        builder.StartTable();
        for (int i = 0; i < 5; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i + 1} Cell 2");

            // End the current row.
            builder.EndRow();
        }
        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Disable breaking across pages for each row.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = "TableRowBreakDisabled.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
