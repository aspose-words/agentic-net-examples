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

        // Enable different headers/footers for first page and odd/even pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        // Insert a table into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        InsertSampleTable(builder, "Header");

        // Insert a table into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        InsertSampleTable(builder, "Footer");

        // Add some body content so the document is not empty.
        builder.MoveToSection(0);
        builder.Writeln("Document body starts here.");
        builder.Writeln("Additional content...");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderFooterTable.docx");
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");

        // Indicate success.
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Helper method that builds a simple 2x2 table at the builder's current position.
    private static void InsertSampleTable(DocumentBuilder builder, string location)
    {
        // Start the table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write($"{location} Table - Row 1, Cell 1");
        builder.InsertCell();
        builder.Write($"{location} Table - Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write($"{location} Table - Row 2, Cell 1");
        builder.InsertCell();
        builder.Write($"{location} Table - Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();
    }
}
