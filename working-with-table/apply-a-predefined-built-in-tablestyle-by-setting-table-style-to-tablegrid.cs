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

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndTable(); // Ends the table and returns the Table node.

            // Apply the built‑in "TableGrid" style to the table.
            // The style can be obtained from the document's style collection.
            Style tableGridStyle = doc.Styles[StyleIdentifier.TableGrid];
            table.Style = tableGridStyle;

            // Save the document to a local file.
            string outputPath = "TableStyleExample.docx";
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // The program finishes without waiting for user input.
        }
    }
}
