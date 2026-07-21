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

        // Move the builder cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start a table inside the header.
        Table table = builder.StartTable();

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row with a single cell that spans two columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // Begin horizontal merge.
        builder.Write("Spanned Cell");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue merge.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Return to the main document body and add a sample paragraph.
        builder.MoveToSection(0);
        builder.Writeln("Document body text.");

        // Save the document to disk.
        string outputPath = "HeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
