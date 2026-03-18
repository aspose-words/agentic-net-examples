using System;
using Aspose.Words;
using Aspose.Words.Tables;

class PreventTableSplitAcrossColumns
{
    static void Main()
    {
        // Load an existing document that contains a multi‑column layout.
        Document doc = new Document("Input.docx");

        // Iterate through all tables in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // For each row in the table disable breaking across columns/pages.
            foreach (Row row in table.Rows)
            {
                // Setting this flag to false keeps the row together when the
                // table is placed in a multi‑column section.
                row.RowFormat.AllowBreakAcrossPages = false;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
