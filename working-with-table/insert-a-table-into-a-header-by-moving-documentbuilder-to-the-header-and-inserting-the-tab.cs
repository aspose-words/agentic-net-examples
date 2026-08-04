using System;
using System.IO;
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

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder's cursor to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Start a new table in the header.
            Table table = builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Header Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Header Cell 2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Header Cell 3");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Header Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Define the output path.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "HeaderTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine("Document saved successfully to: " + outputPath);
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
