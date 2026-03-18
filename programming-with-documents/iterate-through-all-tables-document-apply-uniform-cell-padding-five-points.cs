using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyUniformCellPadding
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Get all tables in the document (including nested tables).
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table.
        foreach (Table table in tables)
        {
            // Iterate through each row of the current table.
            foreach (Row row in table.Rows)
            {
                // Iterate through each cell of the current row.
                foreach (Cell cell in row.Cells)
                {
                    // Apply a uniform padding of 5 points on all sides.
                    cell.CellFormat.SetPaddings(5, 5, 5, 5);
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
