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

        // Set the document to have two columns.
        doc.FirstSection.PageSetup.TextColumns.SetCount(2);

        // Build a simple 2x2 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Prevent the table rows from breaking across columns/pages.
        // The Table class does not expose an AllowBreakAcrossColumns property.
        // Instead, set the RowFormat.AllowBreakAcrossPages property for each row.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Table_NoColumnBreak.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
