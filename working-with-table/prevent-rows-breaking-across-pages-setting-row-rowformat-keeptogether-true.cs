using System;
using Aspose.Words;
using Aspose.Words.Tables;

class PreventRowBreaks
{
    static void Main()
    {
        // Load an existing document (replace with your actual file path)
        Document doc = new Document("Input.docx");

        // Iterate through all tables in the document
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate through each row in the current table
            foreach (Row row in table.Rows)
            {
                // Prevent the row from breaking across pages.
                // Setting AllowBreakAcrossPages to false keeps the entire row together.
                row.RowFormat.AllowBreakAcrossPages = false;
            }
        }

        // Save the modified document (replace with your desired output path)
        doc.Save("Output.docx");
    }
}
