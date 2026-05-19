using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTablePadding
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table as sample content.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            // Iterate through all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Iterate through each row of the current table.
                foreach (Row row in table.Rows)
                {
                    // Iterate through each cell of the current row.
                    foreach (Cell cell in row.Cells)
                    {
                        // Apply a uniform padding of 5 points on all sides of the cell.
                        cell.CellFormat.SetPaddings(5, 5, 5, 5);
                    }
                }
            }

            // Save the modified document.
            doc.Save("PaddedTables.docx");
        }
    }
}
