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

            // Build a sample table with two rows and two columns.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndTable();

            // Iterate through all tables in the document.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                // Iterate through all cells of the current table.
                NodeCollection cells = tbl.GetChildNodes(NodeType.Cell, true);
                foreach (Cell cell in cells)
                {
                    // Apply a uniform padding of 5 points on all sides.
                    cell.CellFormat.SetPaddings(5, 5, 5, 5);
                }
            }

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
