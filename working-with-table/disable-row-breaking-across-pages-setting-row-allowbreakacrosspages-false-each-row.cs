using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document that contains a table spanning multiple pages.
        Document doc = new Document("Table spanning two pages.docx");

        // Find every table in the document (searches all sections).
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and then each row within the table.
        foreach (Table table in tables)
        {
            foreach (Row row in table.Rows)
            {
                // Disable the ability for this row to break across a page boundary.
                row.RowFormat.AllowBreakAcrossPages = false;
            }
        }

        // Save the updated document.
        doc.Save("Table.AllowBreakAcrossPages.docx");
    }
}
