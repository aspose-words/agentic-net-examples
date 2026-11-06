using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with 2 rows and 3 columns.
            builder.StartTable();

            // First row.
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R1C{col + 1}");
            }
            builder.EndRow();

            // Second row.
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R2C{col + 1}");
            }
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Create a new row to be added at the end of the table.
            Row newRow = new Row(doc);

            // Fill each cell of the new row with placeholder text.
            int columnCount = table.Rows[0].Cells.Count;
            for (int i = 0; i < columnCount; i++)
            {
                Cell cell = new Cell(doc);
                // Add a paragraph with placeholder text.
                Paragraph para = new Paragraph(doc);
                para.AppendChild(new Run(doc, $"Placeholder {i + 1}"));
                cell.AppendChild(para);
                newRow.Cells.Add(cell);
            }

            // Append the new row to the table.
            table.Rows.Add(newRow);

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
