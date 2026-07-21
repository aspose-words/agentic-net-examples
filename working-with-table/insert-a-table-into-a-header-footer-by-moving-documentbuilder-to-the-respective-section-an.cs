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
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);

        // Add a primary footer to the first section.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert a table into the header.
        // -------------------------------------------------
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start a 2x2 table.
        Table headerTable = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Header Cell 3");
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // -------------------------------------------------
        // Insert a table into the footer.
        // -------------------------------------------------
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Start a 2x2 table.
        Table footerTable = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Footer Cell 1");
        builder.InsertCell();
        builder.Write("Footer Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Footer Cell 3");
        builder.InsertCell();
        builder.Write("Footer Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "HeaderFooterTable.docx");
        doc.Save(outputPath);
    }
}
