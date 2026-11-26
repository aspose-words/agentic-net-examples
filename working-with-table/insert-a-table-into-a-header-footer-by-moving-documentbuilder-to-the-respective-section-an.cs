using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInHeaderFooterExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------- Insert table into the primary header --------------------
            // Move the builder to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Ensure there is a paragraph to host the table.
            builder.Writeln();

            // Build a simple 2x2 table in the header.
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Header Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Header Cell 2");

            // End first row.
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Header Cell 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Header Cell 4");

            // End second row and the table.
            builder.EndRow();
            builder.EndTable();

            // -------------------- Insert table into the primary footer --------------------
            // Move the builder to the primary footer of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // Ensure there is a paragraph to host the table.
            builder.Writeln();

            // Build a simple 2x2 table in the footer.
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Footer Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Footer Cell 2");

            // End first row.
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Footer Cell 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Footer Cell 4");

            // End second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            string outputPath = "TableInHeaderFooter.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
