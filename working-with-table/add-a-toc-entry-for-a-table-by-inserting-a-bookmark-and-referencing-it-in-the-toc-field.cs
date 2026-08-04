using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark that will surround the table.
        builder.StartBookmark("TableBookmark");

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the bookmark after the table.
        builder.EndBookmark("TableBookmark");

        // Insert a paragraph break before the TOC.
        builder.Writeln();

        // Insert a TOC field that references the bookmark containing the table.
        FieldToc tocField = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
        tocField.BookmarkName = "TableBookmark";

        // Update fields so the TOC reflects the current document structure.
        doc.UpdateFields();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableWithToc.docx");
        doc.Save(outputPath);
    }
}
