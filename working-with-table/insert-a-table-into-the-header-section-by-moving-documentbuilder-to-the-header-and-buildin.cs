using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HeaderTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Start a table in the header.
            Table table = builder.StartTable();

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

            // Return to the main document body (optional) and add some content.
            builder.MoveToSection(0);
            builder.Writeln("Document body content.");

            // Save the document.
            doc.Save("HeaderTable.docx");
        }
    }
}
