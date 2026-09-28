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

        // Add a primary header to the first section.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);

        // Add a primary footer to the first section.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);

        // ----- Insert a table into the header -----
        // Move the builder's cursor to the header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Build a simple 2x2 table in the header.
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

        // ----- Insert a table into the footer -----
        // Move the builder's cursor to the footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Build a simple 1x3 table in the footer.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Write("Footer Cell A");

        // Second cell.
        builder.InsertCell();
        builder.Write("Footer Cell B");

        // Third cell.
        builder.InsertCell();
        builder.Write("Footer Cell C");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "HeaderFooterTable.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
