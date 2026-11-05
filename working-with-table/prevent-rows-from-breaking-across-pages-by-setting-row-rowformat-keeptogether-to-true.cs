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

        // Build a simple table with several rows.
        builder.StartTable();

        // Add 5 rows, each with a single cell containing some text.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i} - This row should stay together on a page.");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Prevent each row from breaking across pages.
        foreach (Row row in table.Rows)
        {
            // Setting AllowBreakAcrossPages to false keeps the row together.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to a file.
        string outputPath = "KeepTogetherRows.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
