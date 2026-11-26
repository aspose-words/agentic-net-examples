using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a custom TOC field that includes only entries with the "Table" caption label.
        builder.InsertField(@"TOC \h \z \c ""Table""");
        builder.Writeln(); // Move to a new paragraph after the TOC.

        // Insert tables with captions.
        InsertTableWithCaption(builder, "First table caption");
        InsertTableWithCaption(builder, "Second table caption");
        InsertTableWithCaption(builder, "Third table caption");

        // Update all fields (TOC, SEQ, etc.) so the TOC reflects the table captions.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "TableOfContentsForTables.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }

    private static void InsertTableWithCaption(DocumentBuilder builder, string captionText)
    {
        // Insert a paragraph containing the SEQ field for table numbering and the caption text.
        builder.InsertField("SEQ Table \\* ARABIC");
        builder.Writeln($": {captionText}");

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Add a blank paragraph after the table for spacing.
        builder.Writeln();
    }
}
