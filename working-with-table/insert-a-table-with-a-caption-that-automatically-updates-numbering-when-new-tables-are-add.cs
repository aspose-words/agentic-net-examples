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

        // Insert first table with a caption.
        InsertTableWithCaption(builder, "First table caption");

        // Insert second table with a caption – the SEQ field will auto‑increment.
        InsertTableWithCaption(builder, "Second table caption");

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablesWithCaptions.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }

    private static void InsertTableWithCaption(DocumentBuilder builder, string captionText)
    {
        // Ensure we are on a new paragraph before the caption.
        builder.Writeln();

        // Insert a SEQ field that automatically numbers tables.
        builder.InsertField("SEQ Table \\* ARABIC");
        builder.Write(" " + captionText);
        builder.Writeln(); // End the caption paragraph.

        // Build a simple 2×2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Add a blank line after the table for readability.
        builder.Writeln();
    }
}
