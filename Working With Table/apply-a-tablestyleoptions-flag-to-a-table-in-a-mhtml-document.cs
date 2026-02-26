using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptionsToMhtml
{
    static void Main()
    {
        // Load the MHTML document.
        // The constructor of Document loads the file based on its format.
        Document doc = new Document("InputDocument.mht");

        // Find the first table in the document.
        // GetChild searches the document tree for a node of the specified type.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply desired style options to the table.
        // Here we combine FirstRow and RowBands as an example.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document back to MHTML format.
        doc.Save("OutputDocument.mht", SaveFormat.Mhtml);
    }
}
