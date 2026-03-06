using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing Word document (WORDML or DOCX).
        Document doc = new Document("Input.docx");

        // Retrieve the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;

        if (table != null)
        {
            // Apply the desired TableStyleOptions flags.
            // Example: apply first column formatting, row banding, and first row formatting.
            table.StyleOptions = TableStyleOptions.FirstColumn |
                                 TableStyleOptions.RowBands |
                                 TableStyleOptions.FirstRow;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
