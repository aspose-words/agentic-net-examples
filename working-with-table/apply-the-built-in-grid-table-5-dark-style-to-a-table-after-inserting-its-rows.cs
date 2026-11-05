using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            builder.StartTable();

            // Insert 3 rows with 3 cells each.
            for (int row = 1; row <= 3; row++)
            {
                for (int col = 1; col <= 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"Row {row} Cell {col}");
                }
                builder.EndRow();
            }

            // End the table construction.
            builder.EndTable();

            // Retrieve the created table from the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Apply the built‑in "Grid Table 5 Dark" style.
            table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

            // Save the document to disk.
            string outputPath = "TableStyleExample.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
