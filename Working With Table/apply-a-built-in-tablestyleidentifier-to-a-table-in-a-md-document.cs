using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load a Markdown document that contains at least one table.
        Document doc = new Document("input.md");

        // Find the first table in the document.
        Table table = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (table != null)
        {
            // Apply a built‑in table style using its identifier.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Optionally, specify which parts of the style to apply.
            table.StyleOptions = TableStyleOptions.FirstRow |
                                 TableStyleOptions.FirstColumn |
                                 TableStyleOptions.RowBands;

            // Adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);
        }

        // Save the modified document to a Word format (DOCX).
        doc.Save("output.docx");
    }
}
