using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoHeader
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start a new table inside the header.
        Table headerTable = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 4");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Return the cursor to the main document body (first section).
        builder.MoveToSection(0);

        // Add some body content to demonstrate the header.
        builder.Writeln("This is the main document body.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Save the document to a file.
        doc.Save("HeaderWithTable.docx");
    }
}
