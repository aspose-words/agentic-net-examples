using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start a table inside the header.
        Table headerTable = builder.StartTable();

        // First row with two cells.
        builder.InsertCell();
        builder.Write("Header Cell 1");
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row with two cells.
        builder.InsertCell();
        builder.Write("Header Cell 3");
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // End the table. The cursor now sits after the table in the header.
        builder.EndTable();

        // Return the cursor to the main body of the first section.
        builder.MoveToSection(0);

        // Add some body text so the header can be seen in the output document.
        builder.Writeln("This is the main document body.");

        // Save the document to a file.
        doc.Save("HeaderTable.docx");
    }
}
