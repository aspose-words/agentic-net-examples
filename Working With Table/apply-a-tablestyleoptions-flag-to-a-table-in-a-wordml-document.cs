using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Load an existing WORDML (WordprocessingML) document.
        // Replace "Input.docx" with the path to your source document.
        Document doc = new Document("Input.docx");

        // Retrieve the first table in the document.
        // The GetChild method searches the document tree for a node of the specified type.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
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

        // Optionally, set a built‑in style identifier so the style options have a base style to work from.
        // This step is not required if the table already has a style assigned.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Save the modified document.
        // The output format can be DOCX, PDF, etc.; here we save as DOCX.
        doc.Save("Output.docx");
    }
}
