using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Try to get the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;

        // If the document does not contain a table, create a simple one.
        if (table == null)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Sample cell");
            builder.EndTable();
        }

        // Apply the desired TableStyleOptions flags.
        // Example: apply formatting to the first row and enable row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
