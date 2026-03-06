using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Path to the folder that contains the DOTX template.
        string dataDir = @"C:\Data\";

        // Load the DOTX document.
        Document doc = new Document(dataDir + "Template.dotx");

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply desired style options (example: first row and row banding).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the modified document.
        doc.Save(dataDir + "Result.docx");
    }
}
