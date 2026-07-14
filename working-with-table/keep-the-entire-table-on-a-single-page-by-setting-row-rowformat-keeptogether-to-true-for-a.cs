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

        // Build a simple 3x2 table.
        Table table = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Prevent rows from breaking across pages.
        // RowFormat does not have a KeepTogether property; the equivalent is AllowBreakAcrossPages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "KeepTableTogether.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
