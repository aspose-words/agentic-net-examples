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

        // Set the section to have two columns (multi‑column layout).
        builder.CurrentSection.PageSetup.TextColumns.SetCount(2);

        // Start a table that would normally flow across the columns.
        Table table = builder.StartTable();

        // Populate the table with a few rows and cells.
        for (int i = 0; i < 3; i++)
        {
            for (int j = 0; j < 2; j++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Cell {j + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Prevent the table rows from breaking across pages/columns.
        // Aspose.Words does not have an AllowBreakAcrossColumns property on Table.
        // Instead, set the RowFormat.AllowBreakAcrossPages property for each row,
        // which also controls breaking across columns in a multi‑column layout.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Table_NoBreakAcrossColumns.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
