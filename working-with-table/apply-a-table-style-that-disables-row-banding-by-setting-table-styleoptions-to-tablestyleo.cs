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
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Disable row banding by ensuring the RowBands flag is not set.
        // Using TableStyleOptions.None removes all conditional formatting, including row banding.
        table.StyleOptions = TableStyleOptions.None;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableNoRowBanding.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The document was not saved correctly.");
        }
    }
}
