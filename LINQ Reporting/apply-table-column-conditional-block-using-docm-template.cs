using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableColumnConditionalExample
{
    static void Main()
    {
        // Load the DOCM template that contains a table.
        Document doc = new Document("Template.docm");

        // Locate the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the template.");
            return;
        }

        // Index of the column we want to conditionally keep or remove (zero‑based).
        int columnIndex = 2; // third column

        // Example condition – replace with your own logic.
        bool keepColumn = false;

        if (!keepColumn)
        {
            // Remove the entire column by deleting the cell at columnIndex in every row.
            // Note: Removing a cell automatically shifts the remaining cells left.
            foreach (Row row in table.Rows)
            {
                // Ensure the row actually has enough cells.
                if (row.Cells.Count > columnIndex)
                {
                    Cell cellToRemove = row.Cells[columnIndex];
                    cellToRemove.Remove();
                }
            }
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
