using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TablePagination
{
    static void Main()
    {
        // Load the source document that contains the large table.
        Document doc = new Document("input.docx");

        // Process every table in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Enable breaking of rows across pages for the entire table.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = true;

                // Allow the row height to be determined automatically.
                // Setting a small minimum height prevents rows from collapsing completely.
                row.RowFormat.HeightRule = HeightRule.Auto;
                row.RowFormat.Height = 12; // points (adjust as needed)
            }
        }

        // Save the updated document.
        doc.Save("output.docx");
    }
}
