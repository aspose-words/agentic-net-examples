using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInsertRowExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3x2 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("R3C1");
            builder.InsertCell();
            builder.Write("R3C2");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Clone the first row to create a new row template.
            Row newRow = (Row)table.Rows[0].Clone(true);

            // Replace cell contents in the new row.
            newRow.Cells[0].Paragraphs[0].Runs.Clear();
            newRow.Cells[0].Paragraphs[0].AppendChild(new Run(doc, "Inserted 1"));
            newRow.Cells[1].Paragraphs[0].Runs.Clear();
            newRow.Cells[1].Paragraphs[0].AppendChild(new Run(doc, "Inserted 2"));

            // Insert the new row after the second row (index 1).
            // The new row should be placed at index 2 in the collection.
            table.Rows.Insert(2, newRow);

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output file was not created.");

            // Simple verification of row count (should be 4).
            if (table.Rows.Count != 4)
                throw new Exception("Row insertion failed; unexpected row count.");
        }
    }
}
