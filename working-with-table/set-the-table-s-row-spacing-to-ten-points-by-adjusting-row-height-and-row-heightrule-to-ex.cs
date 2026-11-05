using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableRowSpacingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            // First row, first cell.
            builder.InsertCell();
            builder.Writeln("Row 1, Cell 1");

            // First row, second cell.
            builder.InsertCell();
            builder.Writeln("Row 1, Cell 2");

            // End first row.
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Writeln("Row 2, Cell 1");

            // Second row, second cell.
            builder.InsertCell();
            builder.Writeln("Row 2, Cell 2");

            // End second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Set each row's height to exactly 10 points.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.Height = 10; // Height in points.
                row.RowFormat.HeightRule = HeightRule.Exactly;
            }

            // Save the document.
            string outputPath = "TableRowSpacing.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
