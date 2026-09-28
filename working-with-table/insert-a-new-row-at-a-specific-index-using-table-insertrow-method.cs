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

            // Build a simple table with two rows and one cell each.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Clone the first row to preserve formatting.
            Row newRow = (Row)table.Rows[0].Clone(true);

            // Change the text of the new row's cell.
            Cell newCell = newRow.Cells[0];
            // Clear existing runs.
            newCell.Paragraphs[0].Runs.Clear();
            // Add new text.
            newCell.Paragraphs[0].AppendChild(new Run(doc, "Inserted Row, Cell 1"));

            // Insert the new row at index 1 (between the original rows).
            table.Rows.Insert(1, newRow);

            // Save the document.
            string outputPath = "InsertedRow.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
