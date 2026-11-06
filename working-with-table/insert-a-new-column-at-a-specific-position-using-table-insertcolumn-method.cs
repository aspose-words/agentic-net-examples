using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableColumnInsertionExample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑row, 3‑column table.
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

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Insert a new column at position 1 (between the first and second columns).
            int insertColumnIndex = 1;
            foreach (Row row in table.Rows)
            {
                // Create a new cell for the current row.
                Cell newCell = new Cell(doc);
                // Add a paragraph so that the cell can contain text.
                Paragraph para = new Paragraph(doc);
                newCell.AppendChild(para);
                // Add sample text to the newly inserted cell.
                para.AppendChild(new Run(doc, "NewCol"));

                // Insert the new cell into the row at the desired column index.
                row.Cells.Insert(insertColumnIndex, newCell);
            }

            // Save the document.
            string outputPath = "InsertedColumn.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }
    }
}
