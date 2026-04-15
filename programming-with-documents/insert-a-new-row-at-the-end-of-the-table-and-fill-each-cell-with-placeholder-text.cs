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

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ----- First existing row -----
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // ----- Second existing row -----
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // ----- Insert new row at the end of the table -----
        // The builder is already positioned after the last EndRow, so we can add a new row directly.
        builder.InsertCell();
        builder.Write("Placeholder 1");
        builder.InsertCell();
        builder.Write("Placeholder 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define output path (in the current directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithNewRow.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
