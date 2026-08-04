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

        // Set the section to have two columns.
        doc.FirstSection.PageSetup.TextColumns.SetCount(2);

        // Build a simple 3x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Prevent the table rows from breaking across columns/pages.
        // The Table class does not expose an AllowBreakAcrossColumns property.
        // Instead, set the RowFormat.AllowBreakAcrossPages property for each row.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table.AllowBreakAcrossColumns.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
