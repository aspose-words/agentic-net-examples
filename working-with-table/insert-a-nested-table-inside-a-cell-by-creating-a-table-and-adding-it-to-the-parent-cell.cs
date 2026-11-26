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

            // Build the outer table with two cells.
            builder.StartTable();

            // First cell of the outer table.
            builder.InsertCell();
            builder.Writeln("Outer Cell 1");

            // Second cell of the outer table – this will contain the nested table.
            builder.InsertCell();
            builder.Writeln("Outer Cell 2 (contains nested table)");

            // Insert the nested table inside the current cell.
            // The builder is already positioned inside the second cell.
            builder.StartTable();

            // First row of the nested table.
            builder.InsertCell();
            builder.Writeln("Nested Cell 1");
            builder.EndRow();

            // Second row of the nested table.
            builder.InsertCell();
            builder.Writeln("Nested Cell 2");
            builder.EndRow();

            // Finish the nested table.
            builder.EndTable();

            // Finish the outer table.
            builder.EndRow();
            builder.EndTable();

            // Save the document.
            string outputPath = "NestedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
