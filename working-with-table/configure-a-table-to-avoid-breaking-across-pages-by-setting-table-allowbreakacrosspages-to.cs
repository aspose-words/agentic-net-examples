using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x1 table.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Configure the table to avoid breaking across pages.
        // The Table class does not have an AllowBreakAcrossPages property.
        // Instead, set the property on each row's RowFormat.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = "TableAllowBreakAcrossPages.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
