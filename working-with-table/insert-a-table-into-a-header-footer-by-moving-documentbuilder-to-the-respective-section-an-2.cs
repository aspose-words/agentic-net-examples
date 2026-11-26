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

        // Ensure the first section has a primary header.
        Section section = doc.FirstSection;
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(header);

        // Move the builder to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Build a 2x2 table inside the header.
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

        // End the table.
        builder.EndTable();

        // Ensure the first section has a primary footer.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        section.HeadersFooters.Add(footer);

        // Move the builder to the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Build a 2x2 table inside the footer.
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

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "HeaderFooterTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optional: indicate success.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
