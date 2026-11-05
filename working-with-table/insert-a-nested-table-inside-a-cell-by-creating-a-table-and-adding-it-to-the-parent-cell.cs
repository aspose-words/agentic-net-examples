using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace NestedTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the outer table (2 rows x 2 columns).
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Writeln("Outer Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Writeln("Outer Cell 2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Writeln("Outer Cell 3");

            // Second row, second cell – this cell will contain the nested table.
            builder.InsertCell();

            // Build the nested table inside the current cell.
            builder.StartTable();

            // Nested table first row, first cell.
            builder.InsertCell();
            builder.Writeln("Nested 1");

            // Nested table first row, second cell.
            builder.InsertCell();
            builder.Writeln("Nested 2");
            builder.EndRow();

            // End the nested table.
            builder.EndTable();

            // End the second row of the outer table.
            builder.EndRow();

            // End the outer table.
            builder.EndTable();

            // Save the document to a file.
            string outputPath = "NestedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file: {outputPath}");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
