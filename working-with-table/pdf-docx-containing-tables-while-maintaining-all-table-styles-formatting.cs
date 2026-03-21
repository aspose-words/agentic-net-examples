using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace AsposeWordsTableToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();

            // Build a simple table with some styled content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document with a table:");
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Font.Bold = true;
            builder.Writeln("Header 1");
            builder.InsertCell();
            builder.Font.Bold = true;
            builder.Writeln("Header 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Font.Bold = false;
            builder.Writeln("Cell A1");
            builder.InsertCell();
            builder.Writeln("Cell A2");
            builder.EndRow();

            builder.EndTable();

            // Convert any formatting defined in table styles to direct formatting.
            // This ensures that the visual appearance of tables is preserved
            // when the document is saved to a fixed‑page format such as PDF.
            doc.ExpandTableStylesToDirectFormatting();

            // Determine output path in the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputDocument.pdf");

            // Save the document as PDF.
            doc.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved to: {outputPath}");
        }
    }
}
