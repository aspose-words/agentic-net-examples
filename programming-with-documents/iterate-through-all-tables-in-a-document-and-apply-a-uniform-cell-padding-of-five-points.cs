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

            // Build a sample table with a few rows and cells.
            Table table = builder.StartTable();
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {i + 1}, Cell {j + 1}");
                }
                builder.EndRow();
            }
            builder.EndTable();

            // Iterate through all tables in the document and set uniform cell padding of 5 points.
            NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                // Apply padding to the whole table; this affects all its cells.
                tbl.LeftPadding = 5;
                tbl.RightPadding = 5;
                tbl.TopPadding = 5;
                tbl.BottomPadding = 5;
            }

            // Save the modified document.
            doc.Save("TablePaddingResult.docx");
        }
    }
}
