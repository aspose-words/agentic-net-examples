using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Ensure the document has at least one section.
        Section section = doc.FirstSection;

        // Add a primary header to the section. The HeaderFooter constructor requires the owning document.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(header);

        // Create a DocumentBuilder and move it to the header.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a paragraph to provide a valid container for the table.
        builder.Writeln();

        // Build a 2x2 table in the header.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 2");

        // End first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 4");

        // End second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputPath = "HeaderTable.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
