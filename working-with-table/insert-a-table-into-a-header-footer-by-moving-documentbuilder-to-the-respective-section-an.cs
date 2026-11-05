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

        // Move the builder to the header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Build a 2x2 table in the header.
        builder.StartTable();

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

        builder.EndTable();

        // Add a primary footer to the first section.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        doc.FirstSection.HeadersFooters.Add(footer);

        // Move the builder to the footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Build a 2x2 table in the footer.
        builder.StartTable();

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

        builder.EndTable();

        // Save the document.
        string outputPath = "HeaderFooterTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
