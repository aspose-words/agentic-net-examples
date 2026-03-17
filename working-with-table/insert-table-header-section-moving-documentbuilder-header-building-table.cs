using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HeaderTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Associate a DocumentBuilder with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder's cursor to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Start building a table inside the header.
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

            // End the second row.
            builder.EndRow();

            // Finish the table. The cursor now points after the table in the header.
            builder.EndTable();

            // Return the cursor to the main body of the first section.
            builder.MoveToSection(0);

            // Add some regular document content.
            builder.Writeln("This is the main body of the document.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page content.");

            // Save the document to disk.
            doc.Save("HeaderTable.docx");
        }
    }
}
