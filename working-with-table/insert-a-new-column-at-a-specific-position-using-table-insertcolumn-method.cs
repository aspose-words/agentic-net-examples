using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace InsertColumnExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x3 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Writeln("R1C1");
            builder.InsertCell();
            builder.Writeln("R1C2");
            builder.InsertCell();
            builder.Writeln("R1C3");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Writeln("R2C1");
            builder.InsertCell();
            builder.Writeln("R2C2");
            builder.InsertCell();
            builder.Writeln("R2C3");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Insert a new column at index 1 (second position) by adding a cell to each row.
            foreach (Row row in table.Rows)
            {
                // Create a new cell and insert it at the desired column index.
                Cell newCell = new Cell(doc);
                row.Cells.Insert(1, newCell);

                // Clear any default content and add the desired text.
                newCell.RemoveAllChildren();
                Paragraph para = new Paragraph(doc);
                newCell.AppendChild(para);
                para.AppendChild(new Run(doc, "NewCol"));
            }

            // Save the document to disk.
            string outputPath = "InsertedColumn.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);

            // Optional validation: ensure column count increased.
            int expectedColumns = 4; // Original 3 + 1 inserted.
            int actualColumns = table.Rows[0].Cells.Count;
            if (actualColumns != expectedColumns)
                throw new InvalidOperationException($"Column insertion failed. Expected {expectedColumns}, but got {actualColumns}.");

            // Program completed successfully.
        }
    }
}
