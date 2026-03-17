using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Get all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Process each table.
        foreach (Table table in tables)
        {
            // Iterate rows in reverse order so that removal does not affect the index of remaining rows.
            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                Row row = table.Rows[i];

                // If the row's height is zero, it is considered empty and should be removed.
                if (row.RowFormat.Height == 0)
                {
                    row.Remove();
                }
            }
        }

        // Save the modified document.
        doc.Save("output.docx");
    }
}
