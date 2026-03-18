using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the cell format globally before any cells are inserted.
        // This setting will be applied to every cell created afterwards.
        builder.CellFormat.Orientation = TextOrientation.Downward; // vertical top‑to‑bottom text

        // Build a sample table (3 rows × 2 columns) to demonstrate the global setting.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // If the document already contains tables, enforce vertical orientation
        // for all existing cells by iterating through them.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
            {
                cell.CellFormat.Orientation = TextOrientation.Downward;
            }
        }

        // Save the resulting document.
        doc.Save("VerticalTable.docx");
    }
}
