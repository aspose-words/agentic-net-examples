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

        // Add a primary header to the first section.
        Section section = doc.FirstSection;
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(header);

        // Create a DocumentBuilder and move it to the header.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Build a simple 2x2 table inside the header.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "HeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
