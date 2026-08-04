using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace HeaderFooterTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Insert a table into the primary header.
            // -------------------------------------------------
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            // Start building the table.
            Table headerTable = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Header Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Header Cell 2");

            // End the first row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // -------------------------------------------------
            // Insert a table into the primary footer.
            // -------------------------------------------------
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            Table footerTable = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Footer Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Footer Cell 2");

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // -------------------------------------------------
            // Return to the main document body and add some text.
            // -------------------------------------------------
            builder.MoveToSection(0);
            builder.Writeln("This is the main body of the document.");

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderFooterTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // Inform the user (no interactive prompts required).
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
