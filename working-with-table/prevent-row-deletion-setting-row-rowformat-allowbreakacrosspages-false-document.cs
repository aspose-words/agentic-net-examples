using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Iterate over every table in the document.
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            // Iterate over each row in the current table.
            foreach (Row row in table.Rows)
            {
                // Prevent the row from breaking across pages.
                row.RowFormat.AllowBreakAcrossPages = false;
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
