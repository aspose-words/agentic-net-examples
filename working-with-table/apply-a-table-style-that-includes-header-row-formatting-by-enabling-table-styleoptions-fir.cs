using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleFirstRowExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Insert the header row cells.
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            // Finish the header row.
            builder.EndRow();

            // Insert a few data rows.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Apply a built‑in table style that supports header rows.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable the FirstRow style option so the first row is formatted as a header.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Optionally auto‑fit the table to its contents.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Define the output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableStyleFirstRow.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
