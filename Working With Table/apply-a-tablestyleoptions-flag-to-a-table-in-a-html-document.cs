using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load the HTML document.
        Document doc = new Document("input.html");

        // Find the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply desired style options to the table.
        // Example: apply first row formatting, row banding, and first column formatting.
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.RowBands |
                              TableStyleOptions.FirstColumn;

        // Save the modified document.
        doc.Save("output.docx");
    }
}
